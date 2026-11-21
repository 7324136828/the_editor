//! office-backend: local HTTP service for office-chrome-extension.
//!
//! API (base http://127.0.0.1:47831, see README / CONTRACT):
//!   GET  /health      -> {ok, name, version, api:2, parse:[fmt], render:{word:[..], excel:[..]}, limits}
//!   POST /api/parse   body = file bytes, header X-File-Name = encodeURIComponent(name)
//!                     -> {kind, format, model, warnings:[string], name}
//!   POST /api/render  JSON {kind, model, format, name} -> file bytes with Content-Type,
//!                     X-Output-Format and Content-Disposition (ASCII filename + RFC 5987 filename*)
//!
//! Errors are ALWAYS JSON {"error": "..."} with a meaningful status
//! (400 bad request, 403 forbidden origin/host, 404, 405, 413 too large,
//! 415 unsupported file type, 422 unreadable document / invalid model,
//! 500 internal error — including caught panics; the server keeps running).
//!
//! Security model: the service only listens on loopback and only answers
//! requests that (a) are addressed to a loopback Host name (defeats DNS
//! rebinding) and (b) carry either no Origin (curl, tests, native tools) or
//! a browser-extension Origin. Web pages always send an Origin on cross-site
//! POSTs (including "no-cors" simple requests), so they are rejected with 403
//! before any parser runs. CORS headers are only ever reflected for
//! extension origins.

mod detect;
mod docx;
mod latex;
mod limits;
mod model;
mod pptx;
mod textfmt;
mod xlsx;

use std::net::{IpAddr, Ipv4Addr, SocketAddr};
use std::sync::OnceLock;
use std::time::{Duration, Instant};

use axum::body::{Body, Bytes};
use axum::extract::{DefaultBodyLimit, Request};
use axum::http::{header, HeaderMap, HeaderName, HeaderValue, Method, StatusCode};
use axum::middleware::{self, Next};
use axum::response::{IntoResponse, Response};
use axum::routing::{get, post};
use axum::{Json, Router};
use model::{Block, ExcelBook, LatexDoc, Run, SlidesDoc, WordDoc};
use serde::Deserialize;
use serde_json::{json, Map, Value};
use tower_http::catch_panic::CatchPanicLayer;
use tower_http::cors::{AllowOrigin, CorsLayer};

const VERSION: &str = env!("CARGO_PKG_VERSION");
const API_VERSION: u32 = 2;
const DEFAULT_PORT: u16 = 47831;
const MAX_BODY: usize = limits::MAX_BODY;
/// Output formats /api/render supports, per kind (advertised by /health).
const RENDER_WORD: &[&str] = &["docx", "rtf", "txt", "md", "html", "tex"];
const RENDER_EXCEL: &[&str] = &["xlsx", "csv", "tsv"];
const RENDER_SLIDES: &[&str] = &["pptx", "md", "txt"];
/// LaTeX outputs that need no external tools; "pdf" is added at runtime when
/// a TeX engine is installed (see render_latex_formats).
const RENDER_LATEX: &[&str] = &["tex", "docx", "html", "md", "txt"];

fn render_latex_formats() -> Vec<&'static str> {
    let mut v = RENDER_LATEX.to_vec();
    if latex::compiler().is_some() { v.insert(0, "pdf"); }
    v
}
/// Max parse/render jobs running at once (bounds peak memory).
const MAX_CONCURRENT_JOBS: usize = 4;

const X_FILE_NAME: HeaderName = HeaderName::from_static("x-file-name");
const X_OUTPUT_FORMAT: HeaderName = HeaderName::from_static("x-output-format");
/// Optional render notes: encodeURIComponent(JSON array of strings).
const X_OFFICE_WARNINGS: HeaderName = HeaderName::from_static("x-office-warnings");

// =====================================================================
// Startup / CLI
// =====================================================================

const USAGE: &str = "\
office-backend — local document service for the Office Viewer & Editor extension

USAGE:
    office-backend [--port PORT] [--addr HOST:PORT]

OPTIONS:
    -p, --port PORT       Port to listen on (default 47831, address 127.0.0.1)
    -a, --addr HOST:PORT  Full listen address, e.g. 127.0.0.1:47831 (a bare port also works)
    -h, --help            Show this help
    -V, --version         Show the version

ENVIRONMENT:
    OFFICE_BACKEND_ADDR   Same as --addr (command-line flags win)

After starting, open the extension's Options page and set \"Backend URL\" to the
URL printed at startup (default http://127.0.0.1:47831), then click
\"Test connection\". Only the browser extension can use this service: requests
from web pages are refused.";

#[derive(Debug, PartialEq, Eq)]
enum Cli {
    Run(SocketAddr),
    Help,
    Version,
}

/// Parse "47831", "127.0.0.1:47831", "localhost:47831", "[::1]:47831",
/// "localhost" or "127.0.0.1" into a socket address.
fn parse_addr(s: &str) -> Result<SocketAddr, String> {
    let s = s.trim();
    let bad = || format!("invalid address {s:?} (expected e.g. 127.0.0.1:{DEFAULT_PORT} or just a port number)");
    if s.is_empty() { return Err(bad()); }
    if s.chars().all(|c| c.is_ascii_digit()) {
        let port: u16 = s.parse().map_err(|_| bad())?;
        return Ok(SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), port));
    }
    if let Ok(a) = s.parse::<SocketAddr>() { return Ok(a); }
    if let Ok(ip) = s.trim_start_matches('[').trim_end_matches(']').parse::<IpAddr>() {
        return Ok(SocketAddr::new(ip, DEFAULT_PORT));
    }
    let (host, port) = match s.rsplit_once(':') {
        Some((h, p)) => (h, p.parse::<u16>().map_err(|_| bad())?),
        None => (s, DEFAULT_PORT),
    };
    if host.eq_ignore_ascii_case("localhost") {
        return Ok(SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), port));
    }
    Err(bad())
}

fn parse_cli(args: &[String], env_addr: Option<&str>) -> Result<Cli, String> {
    let mut addr: Option<SocketAddr> = None;
    let mut port: Option<u16> = None;
    let mut it = args.iter();
    while let Some(a) = it.next() {
        let (flag, inline) = match a.split_once('=') {
            Some((f, v)) if f.starts_with("--") => (f, Some(v.to_string())),
            _ => (a.as_str(), None),
        };
        let mut value = |name: &str| -> Result<String, String> {
            inline.clone().or_else(|| it.next().cloned()).ok_or_else(|| format!("{name} needs a value"))
        };
        match flag {
            "-h" | "--help" | "/?" => return Ok(Cli::Help),
            "-V" | "--version" => return Ok(Cli::Version),
            "-p" | "--port" => {
                let v = value("--port")?;
                port = Some(v.trim().parse().map_err(|_| format!("invalid port {v:?}"))?);
            }
            "-a" | "--addr" => addr = Some(parse_addr(&value("--addr")?)?),
            other => return Err(format!("unknown argument {other:?}")),
        }
    }
    let mut a = match addr {
        Some(a) => a,
        None => match env_addr.map(str::trim).filter(|s| !s.is_empty()) {
            Some(e) => parse_addr(e).map_err(|m| format!("OFFICE_BACKEND_ADDR: {m}"))?,
            None => SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), DEFAULT_PORT),
        },
    };
    if let Some(p) = port { a.set_port(p); }
    Ok(Cli::Run(a))
}

fn friendly_bind_error(addr: SocketAddr, e: &std::io::Error) -> String {
    let port = addr.port();
    let next = port.wrapping_add(1).max(1024);
    match e.kind() {
        std::io::ErrorKind::AddrInUse => format!(
            "Port {port} is already in use — is office-backend already running?\n\
             Check http://127.0.0.1:{port}/health in your browser. To use another port run:\n    \
             office-backend --port {next}\n\
             and set \"Backend URL\" in the extension Options to http://127.0.0.1:{next}"),
        std::io::ErrorKind::PermissionDenied => format!(
            "Not allowed to listen on {addr} ({e}). The port may be reserved by the system; try another one, e.g.\n    \
             office-backend --port {next}"),
        std::io::ErrorKind::AddrNotAvailable => format!(
            "The address {addr} is not available on this machine ({e}). Use 127.0.0.1 (the default)."),
        _ => format!("Could not listen on {addr}: {e}"),
    }
}

#[tokio::main]
async fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let env_addr = std::env::var("OFFICE_BACKEND_ADDR").ok();
    let addr = match parse_cli(&args, env_addr.as_deref()) {
        Ok(Cli::Run(a)) => a,
        Ok(Cli::Help) => { println!("{USAGE}"); return; }
        Ok(Cli::Version) => { println!("office-backend {VERSION}"); return; }
        Err(e) => {
            eprintln!("error: {e}\n\n{USAGE}");
            std::process::exit(2);
        }
    };

    // One readable line per panic (the request still gets a JSON 500).
    std::panic::set_hook(Box::new(|info| eprintln!("internal error (recovered): {info}")));

    let listener = match tokio::net::TcpListener::bind(addr).await {
        Ok(l) => l,
        Err(e) => {
            eprintln!("error: {}", friendly_bind_error(addr, &e));
            std::process::exit(1);
        }
    };
    let local = listener.local_addr().unwrap_or(addr);
    let url = if local.ip().is_unspecified() || local.ip().is_loopback() {
        format!("http://127.0.0.1:{}", local.port())
    } else {
        format!("http://{local}")
    };
    println!("office-backend {VERSION} listening on {url}");
    println!("In the extension's Options page, set \"Backend URL\" to: {url}");
    if !local.ip().is_loopback() {
        println!("note: listening on a non-loopback address; requests are still only accepted for Host 127.0.0.1/localhost.");
    }
    println!("Press Ctrl+C to stop.");

    let shutdown = async {
        let _ = tokio::signal::ctrl_c().await;
        println!("Shutting down.");
    };
    if let Err(e) = axum::serve(listener, app()).with_graceful_shutdown(shutdown).await {
        eprintln!("error: server stopped: {e}");
        std::process::exit(1);
    }
}

/// The full application with all middleware (also used by the tests).
fn app() -> Router {
    Router::new()
        .route("/health", get(health))
        .route("/api/parse", post(api_parse))
        .route("/api/render", post(api_render))
        .fallback(not_found)
        // innermost .. outermost
        .layer(DefaultBodyLimit::max(MAX_BODY))
        .layer(CatchPanicLayer::custom(panic_response))
        .layer(middleware::from_fn(json_errors))
        .layer(cors_layer())
        .layer(middleware::from_fn(guard))
        .layer(middleware::from_fn(log_requests))
}

// =====================================================================
// Errors
// =====================================================================

/// An API error: always rendered as JSON {"error": msg}.
#[derive(Debug)]
struct ApiError {
    status: StatusCode,
    msg: String,
}

impl ApiError {
    fn new(status: StatusCode, msg: impl Into<String>) -> Self { ApiError { status, msg: msg.into() } }
    fn bad_request(msg: impl Into<String>) -> Self { Self::new(StatusCode::BAD_REQUEST, msg) }
    fn unprocessable(msg: impl Into<String>) -> Self { Self::new(StatusCode::UNPROCESSABLE_ENTITY, msg) }
    fn internal(msg: impl Into<String>) -> Self { Self::new(StatusCode::INTERNAL_SERVER_ERROR, msg) }
}

/// Short note attached to responses for the request log.
#[derive(Clone)]
struct LogNote(String);

impl IntoResponse for ApiError {
    fn into_response(self) -> Response {
        let mut r = (self.status, Json(json!({ "error": self.msg }))).into_response();
        r.extensions_mut().insert(LogNote(self.msg));
        r
    }
}

async fn not_found(method: Method, uri: axum::http::Uri) -> ApiError {
    ApiError::new(StatusCode::NOT_FOUND, format!("No such endpoint: {method} {}", uri.path()))
}

fn panic_message(p: &(dyn std::any::Any + Send)) -> String {
    p.downcast_ref::<String>().cloned()
        .or_else(|| p.downcast_ref::<&str>().map(|s| s.to_string()))
        .unwrap_or_else(|| "unknown panic".into())
}

fn panic_response(p: Box<dyn std::any::Any + Send + 'static>) -> Response {
    ApiError::internal(format!("Internal error while processing the document: {}", panic_message(&*p))).into_response()
}

/// Turn every non-JSON error response (axum extractor rejections, body
/// limit, 405, ...) into JSON {"error": ...}.
async fn json_errors(req: Request, next: Next) -> Response {
    let resp = next.run(req).await;
    let st = resp.status();
    if !(st.is_client_error() || st.is_server_error()) { return resp; }
    let is_json = resp.headers().get(header::CONTENT_TYPE)
        .and_then(|v| v.to_str().ok())
        .is_some_and(|ct| ct.starts_with("application/json"));
    if is_json { return resp; }
    let (parts, body) = resp.into_parts();
    let text = axum::body::to_bytes(body, 16 * 1024).await.unwrap_or_default();
    let text = String::from_utf8_lossy(&text).trim().to_string();
    let msg = match st {
        StatusCode::PAYLOAD_TOO_LARGE => format!("The file is too large (the limit is {} MB).", MAX_BODY / (1024 * 1024)),
        StatusCode::METHOD_NOT_ALLOWED => "Method not allowed for this endpoint.".to_string(),
        _ if text.is_empty() => st.canonical_reason().unwrap_or("Error").to_string(),
        _ => text,
    };
    let mut out = ApiError::new(st, msg).into_response();
    for (k, v) in parts.headers.iter() {
        if k != header::CONTENT_TYPE && k != header::CONTENT_LENGTH { out.headers_mut().append(k, v.clone()); }
    }
    out
}

// =====================================================================
// Security: Host / Origin guard + CORS
// =====================================================================

/// Host header must name loopback: 127.0.0.1, localhost or [::1], any port.
fn host_allowed(host: &[u8]) -> bool {
    let Ok(h) = std::str::from_utf8(host) else { return false };
    let h = h.trim().to_ascii_lowercase();
    let (name, port) = if h.starts_with('[') {
        match h.find(']') { Some(i) => (&h[..=i], &h[i + 1..]), None => return false }
    } else {
        match h.find(':') { Some(i) => (&h[..i], &h[i..]), None => (h.as_str(), "") }
    };
    let port_ok = port.is_empty()
        || (port.len() >= 2 && port.len() <= 6 && port.starts_with(':') && port[1..].bytes().all(|b| b.is_ascii_digit()));
    port_ok && matches!(name, "127.0.0.1" | "localhost" | "[::1]")
}

/// Only browser-extension origins may call the service.
fn origin_allowed(origin: &[u8]) -> bool {
    let Ok(o) = std::str::from_utf8(origin) else { return false };
    let o = o.trim().to_ascii_lowercase();
    ["chrome-extension://", "extension://", "moz-extension://"].iter().any(|p| {
        o.strip_prefix(p).is_some_and(|id| !id.is_empty() && id.chars().all(|c| c.is_ascii_alphanumeric() || matches!(c, '-' | '.' | '_')))
    })
}

async fn guard(req: Request, next: Next) -> Response {
    // HTTP/2 carries the authority in the URI instead of a Host header.
    let host = req.headers().get(header::HOST).map(|h| h.as_bytes().to_vec())
        .or_else(|| req.uri().authority().map(|a| a.as_str().as_bytes().to_vec()));
    if let Some(h) = host {
        if !host_allowed(&h) {
            return ApiError::new(StatusCode::FORBIDDEN,
                format!("Forbidden: requests must be addressed to 127.0.0.1 or localhost (got Host {:?}).",
                    String::from_utf8_lossy(&h))).into_response();
        }
    }
    if let Some(o) = req.headers().get(header::ORIGIN) {
        if !origin_allowed(o.as_bytes()) {
            return ApiError::new(StatusCode::FORBIDDEN,
                format!("Forbidden: only the browser extension may use this service (Origin {:?} refused).",
                    String::from_utf8_lossy(o.as_bytes()))).into_response();
        }
    } else if req.headers().get("sec-fetch-site")
        .is_some_and(|v| matches!(v.as_bytes(), b"cross-site" | b"same-site"))
    {
        // No Origin but issued by another web page: an <img>/<script>/form
        // GET used for probing. (Extension fetches carry their Origin; a user
        // opening /health in a tab sends Sec-Fetch-Site: none.)
        return ApiError::new(StatusCode::FORBIDDEN,
            "Forbidden: only the browser extension may use this service.").into_response();
    }
    next.run(req).await
}

fn cors_layer() -> CorsLayer {
    CorsLayer::new()
        .allow_origin(AllowOrigin::predicate(|o: &HeaderValue, _| origin_allowed(o.as_bytes())))
        .allow_methods([Method::GET, Method::POST, Method::OPTIONS])
        .allow_headers([header::CONTENT_TYPE, X_FILE_NAME])
        .expose_headers([header::CONTENT_DISPOSITION, X_OUTPUT_FORMAT, X_OFFICE_WARNINGS])
        .allow_private_network(true)
        .max_age(Duration::from_secs(600))
}

// =====================================================================
// Logging
// =====================================================================

/// One line per request: METHOD PATH STATUS MS [note]
async fn log_requests(req: Request, next: Next) -> Response {
    let start = Instant::now();
    let method = req.method().clone();
    let path = req.uri().path().to_string();
    let resp = next.run(req).await;
    let note = resp.extensions().get::<LogNote>().map(|n| format!("  {}", n.0)).unwrap_or_default();
    println!("{method} {path} {} {}ms{note}", resp.status().as_u16(), start.elapsed().as_millis());
    resp
}

fn human_size(n: usize) -> String {
    if n >= 1024 * 1024 { format!("{:.1} MB", n as f64 / (1024.0 * 1024.0)) }
    else if n >= 1024 { format!("{:.1} KB", n as f64 / 1024.0) }
    else { format!("{n} B") }
}

// =====================================================================
// Blocking work
// =====================================================================

fn job_slots() -> &'static tokio::sync::Semaphore {
    static SLOTS: OnceLock<tokio::sync::Semaphore> = OnceLock::new();
    SLOTS.get_or_init(|| tokio::sync::Semaphore::new(MAX_CONCURRENT_JOBS))
}

/// Run CPU-heavy work off the async runtime, with a concurrency cap.
/// Panics inside `f` become a JSON 500 (the server keeps running).
async fn run_blocking<T, F>(f: F) -> Result<T, ApiError>
where
    T: Send + 'static,
    F: FnOnce() -> Result<T, ApiError> + Send + 'static,
{
    let _permit = job_slots().acquire().await.map_err(|_| ApiError::internal("server is shutting down"))?;
    match tokio::task::spawn_blocking(f).await {
        Ok(r) => r,
        Err(e) if e.is_panic() => Err(ApiError::internal(format!(
            "Internal error while processing the document: {}", panic_message(&*e.into_panic())))),
        Err(e) => Err(ApiError::internal(format!("task failed: {e}"))),
    }
}

// =====================================================================
// /health
// =====================================================================

async fn health() -> Json<Value> {
    Json(json!({
        "ok": true,
        "name": "office-backend",
        "version": VERSION,
        "api": API_VERSION,
        "parse": detect::PARSE_FORMATS,
        "render": { "word": RENDER_WORD, "excel": RENDER_EXCEL, "slides": RENDER_SLIDES, "latex": render_latex_formats() },
        "latex": { "compiler": latex::compiler() },
        "limits": { "maxUploadBytes": MAX_BODY },
    }))
}

// =====================================================================
// /api/parse
// =====================================================================

/// Decode the X-File-Name header: percent-decoding (the client sends
/// encodeURIComponent(name)); invalid %-sequences are kept literally;
/// non-UTF-8 results (raw Latin-1 header bytes, or %-encoded Latin-1)
/// are read as Latin-1 so the extension always survives.
fn decode_file_name(raw: &[u8]) -> String {
    fn hex(b: u8) -> Option<u8> {
        match b { b'0'..=b'9' => Some(b - b'0'), b'a'..=b'f' => Some(b - b'a' + 10), b'A'..=b'F' => Some(b - b'A' + 10), _ => None }
    }
    let mut out = Vec::with_capacity(raw.len());
    let mut i = 0;
    while i < raw.len() {
        if raw[i] == b'%' {
            if let (Some(h), Some(l)) = (raw.get(i + 1).copied().and_then(hex), raw.get(i + 2).copied().and_then(hex)) {
                out.push(h * 16 + l);
                i += 3;
                continue;
            }
        }
        out.push(raw[i]);
        i += 1;
    }
    let s = match String::from_utf8(out) {
        Ok(s) => s,
        Err(e) => e.into_bytes().iter().map(|&b| b as char).collect(),
    };
    s.chars().filter(|c| !c.is_control()).collect::<String>().trim().to_string()
}

/// Excel-safe sheet name from a file name ("Q3 sales.csv" -> "Q3 sales").
fn sheet_name_for(file_name: &str) -> String {
    let base = file_name.rsplit(['/', '\\']).next().unwrap_or(file_name);
    let stem = match base.rsplit_once('.') { Some((s, _)) if !s.is_empty() => s, _ => base };
    let cleaned: String = stem.chars()
        .map(|c| if matches!(c, '[' | ']' | ':' | '*' | '?' | '/' | '\\') { '_' } else { c })
        .collect::<String>().trim().trim_matches('\'').chars().take(31).collect();
    if cleaned.trim().is_empty() { "Sheet1".into() } else { cleaned }
}

struct Parsed {
    kind: &'static str,
    format: &'static str,
    model: Value,
    warnings: Vec<String>,
}

fn push_unique(list: &mut Vec<String>, w: String) {
    let w = w.trim().to_string();
    if !w.is_empty() && !list.contains(&w) { list.push(w); }
}

/// Serialize a word model, lifting meta.warnings to the response warnings.
fn word_json(mut doc: WordDoc, fmt: &str, warnings: &mut Vec<String>) -> Result<Value, ApiError> {
    let before = warnings.len();
    match doc.meta.remove("warnings") {
        Some(Value::Array(ws)) => for w in ws {
            push_unique(warnings, match w { Value::String(s) => s, other => other.to_string() });
        },
        Some(Value::String(s)) => push_unique(warnings, s),
        _ => {}
    }
    let lossy = doc.meta.get("lossy").and_then(Value::as_bool).unwrap_or(false);
    if warnings.len() == before && (fmt == "doc" || lossy) {
        let note = doc.meta.get("note").and_then(Value::as_str).map(|s| format!("{} — ", s.trim_end_matches('.'))).unwrap_or_default();
        push_unique(warnings, if fmt == "doc" {
            format!("{note}Legacy Word 97-2003 (.doc) file: only the text could be recovered. Formatting, tables, images and lists are lost; check the content before saving over the original.")
        } else {
            format!("{note}Some content could not be imported exactly.")
        });
    }
    serde_json::to_value(doc).map_err(|e| ApiError::internal(e.to_string()))
}

/// Serialize a slides model, lifting meta.warnings to the response warnings.
fn slides_json(mut deck: SlidesDoc, warnings: &mut Vec<String>) -> Result<Value, ApiError> {
    match deck.meta.remove("warnings") {
        Some(Value::Array(ws)) => for w in ws {
            push_unique(warnings, match w { Value::String(s) => s, other => other.to_string() });
        },
        Some(Value::String(s)) => push_unique(warnings, s),
        _ => {}
    }
    serde_json::to_value(deck).map_err(|e| ApiError::internal(e.to_string()))
}

/// Serialize an excel model, lifting book.warnings to the response warnings.
fn excel_json(mut book: ExcelBook, warnings: &mut Vec<String>) -> Result<Value, ApiError> {
    for w in std::mem::take(&mut book.warnings) { push_unique(warnings, w); }
    serde_json::to_value(book).map_err(|e| ApiError::internal(e.to_string()))
}

/// Detect + parse an upload (runs on a blocking thread).
fn parse_upload(name: &str, bytes: &[u8]) -> Result<Parsed, ApiError> {
    // Zip-bomb / size checks before anything inflates the container
    // (format sniffing itself opens the ZIP directory).
    if detect::sniff(bytes) == Some("zip") {
        limits::check_zip_detailed(bytes).map_err(|e| match e {
            limits::ZipCheckError::TooLarge(m) => ApiError::new(StatusCode::PAYLOAD_TOO_LARGE, m),
            limits::ZipCheckError::Corrupt(m) => ApiError::unprocessable(m),
        })?;
    }
    let det = detect::detect(name, bytes).map_err(|e| {
        // A broken file of a supported type is unprocessable (422); a type we
        // do not handle at all is unsupported (415).
        let status = if e.damaged { StatusCode::UNPROCESSABLE_ENTITY } else { StatusCode::UNSUPPORTED_MEDIA_TYPE };
        ApiError::new(status, e.message)
    })?;
    let fmt = det.format;
    let kind = detect::kind_of_format(fmt).ok_or_else(|| ApiError::internal(format!("no kind for format {fmt}")))?;
    let mut warnings = Vec::new();
    if let Some(n) = det.note { push_unique(&mut warnings, n); }
    let unreadable = |e: String| ApiError::unprocessable(format!("Could not read {}: {e}", detect::describe(fmt)));

    let model = match fmt {
        "docx" => word_json(docx::parse_docx(bytes).map_err(unreadable)?, fmt, &mut warnings)?,
        "odt" => word_json(textfmt::parse_odt(bytes).map_err(unreadable)?, fmt, &mut warnings)?,
        "rtf" => word_json(textfmt::parse_rtf(bytes).map_err(unreadable)?, fmt, &mut warnings)?,
        "txt" => word_json(textfmt::parse_text(bytes, false), fmt, &mut warnings)?,
        "md" => word_json(textfmt::parse_text(bytes, true), fmt, &mut warnings)?,
        "doc" => word_json(textfmt::parse_doc(bytes), fmt, &mut warnings)?,
        "xlsx" | "xlsm" | "xls" | "ods" => excel_json(xlsx::parse_book(bytes, None).map_err(unreadable)?, &mut warnings)?,
        "csv" => excel_json(xlsx::parse_csv_bytes(bytes, None, &sheet_name_for(name)), &mut warnings)?,
        "tsv" => excel_json(xlsx::parse_csv_bytes(bytes, Some('\t'), &sheet_name_for(name)), &mut warnings)?,
        "pptx" => slides_json(pptx::parse_pptx(bytes).map_err(unreadable)?, &mut warnings)?,
        "tex" => {
            let (text, enc_note) = textfmt::decode_text(bytes);
            if let Some(n) = enc_note { push_unique(&mut warnings, n.to_string()); }
            serde_json::to_value(LatexDoc { kind: "latex".into(), source: text }).map_err(|e| ApiError::internal(e.to_string()))?
        }
        other => return Err(ApiError::new(StatusCode::UNSUPPORTED_MEDIA_TYPE, format!("Unsupported format {other}"))),
    };
    Ok(Parsed { kind, format: fmt, model, warnings })
}

async fn api_parse(headers: HeaderMap, body: Bytes) -> Result<Response, ApiError> {
    let name = headers.get(X_FILE_NAME)
        .map(|v| decode_file_name(v.as_bytes()))
        .filter(|s| !s.is_empty())
        .unwrap_or_else(|| "document".into());
    if body.is_empty() {
        return Err(ApiError::bad_request("The file is empty (no request body)."));
    }
    let size = body.len();
    let name2 = name.clone();
    let p = run_blocking(move || parse_upload(&name2, &body)).await?;
    let mut resp = Json(json!({
        "kind": p.kind,
        "format": p.format,
        "model": p.model,
        "warnings": p.warnings,
        "name": name,
    })).into_response();
    let warn = if p.warnings.is_empty() { String::new() } else { format!(", {} warning(s)", p.warnings.len()) };
    resp.extensions_mut().insert(LogNote(format!("{} {}{warn}", p.format, human_size(size))));
    Ok(resp)
}

// =====================================================================
// /api/render
// =====================================================================

#[derive(Deserialize)]
struct RenderReq {
    kind: String,
    model: Value,
    format: String,
    #[serde(default)]
    name: Option<String>,
}

fn content_type_of(format: &str) -> &'static str {
    match format {
        "docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "rtf" => "application/rtf",
        "txt" => "text/plain; charset=utf-8",
        "md" => "text/markdown; charset=utf-8",
        "html" => "text/html; charset=utf-8",
        "xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "csv" => "text/csv; charset=utf-8",
        "tsv" => "text/tab-separated-values; charset=utf-8",
        "pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "tex" => "application/x-tex; charset=utf-8",
        "pdf" => "application/pdf",
        _ => "application/octet-stream",
    }
}

/// Unicode-preserving file name for the download: last path component,
/// characters Windows forbids replaced, known document extension replaced
/// by the output extension.
fn output_filename(name: Option<&str>, ext: &str) -> String {
    let raw = name.unwrap_or("").trim();
    let base = raw.rsplit(['/', '\\']).next().unwrap_or(raw);
    let mut stem = base.to_string();
    if let Some((s, e)) = base.rsplit_once('.') {
        let e = e.to_ascii_lowercase();
        let known = detect::format_of_name(base).is_some()
            || RENDER_WORD.contains(&e.as_str()) || RENDER_EXCEL.contains(&e.as_str())
            || RENDER_SLIDES.contains(&e.as_str()) || RENDER_LATEX.contains(&e.as_str())
            || matches!(e.as_str(), "htm" | "markdown" | "docm" | "dotx" | "xltx" | "xlsb" | "pdf" | "json");
        if known { stem = s.to_string(); }
    }
    let cleaned: String = stem.chars()
        .map(|c| if c.is_control() || matches!(c, '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*') { '_' } else { c })
        .collect();
    let cleaned = cleaned.trim().trim_matches('.').trim();
    let cleaned: String = cleaned.chars().take(150).collect();
    let stem = if cleaned.is_empty() { "document".to_string() } else { cleaned };
    format!("{stem}.{ext}")
}

/// ASCII stand-in for a Unicode file name (filename= fallback): common
/// Latin accents are folded, everything else becomes '_'.
fn ascii_fallback(name: &str) -> String {
    let mut out = String::with_capacity(name.len());
    for c in name.chars() {
        let folded: &str = match c {
            'À'..='Å' => "A", 'à'..='å' => "a", 'Æ' => "AE", 'æ' => "ae", 'Ç' => "C", 'ç' => "c",
            'È'..='Ë' => "E", 'è'..='ë' => "e", 'Ì'..='Ï' => "I", 'ì'..='ï' => "i", 'Ð' => "D", 'ð' => "d",
            'Ñ' => "N", 'ñ' => "n", 'Ò'..='Ö' | 'Ø' => "O", 'ò'..='ö' | 'ø' => "o", 'Ù'..='Ü' => "U", 'ù'..='ü' => "u",
            'Ý' => "Y", 'ý' | 'ÿ' => "y", 'ß' => "ss", 'Œ' => "OE", 'œ' => "oe", 'Š' => "S", 'š' => "s", 'Ž' => "Z", 'ž' => "z",
            '\u{2018}' | '\u{2019}' => "'", '\u{201C}' | '\u{201D}' => "_", '\u{2013}' | '\u{2014}' => "-",
            _ => "",
        };
        if !folded.is_empty() { out.push_str(folded); }
        else if c.is_ascii_graphic() && !matches!(c, '"' | '\\' | '%' | ';') || c == ' ' { out.push(c); }
        else { out.push('_'); }
    }
    out
}

/// RFC 5987 / 8187 attr-char percent-encoding of UTF-8.
fn rfc5987_encode(s: &str) -> String {
    let mut out = String::new();
    for b in s.bytes() {
        if b.is_ascii_alphanumeric() || b"!#$&+-.^_`|~".contains(&b) { out.push(b as char); }
        else { out.push_str(&format!("%{b:02X}")); }
    }
    out
}

/// `attachment; filename="<ascii>"; filename*=UTF-8''<pct>` — always a
/// valid header value.
fn content_disposition(filename: &str) -> HeaderValue {
    let v = format!("attachment; filename=\"{}\"; filename*=UTF-8''{}", ascii_fallback(filename), rfc5987_encode(filename));
    HeaderValue::from_str(&v).unwrap_or_else(|_| HeaderValue::from_static("attachment; filename=\"document\""))
}

/// encodeURIComponent-compatible encoding (for X-Office-Warnings).
fn uri_component_encode(s: &str) -> String {
    let mut out = String::new();
    for b in s.bytes() {
        if b.is_ascii_alphanumeric() || b"-_.!~*'()".contains(&b) { out.push(b as char); }
        else { out.push_str(&format!("%{b:02X}")); }
    }
    out
}

// ---------- word model normalisation (no silent block loss) ----------

const BLOCK_TYPES_WITH_RUNS: &[&str] = &["p", "h1", "h2", "h3", "h4", "h5", "h6", "quote", "code"];

fn coerce_uint(v: &Value) -> Option<u64> {
    let f = match v {
        Value::Number(n) => n.as_f64()?,
        Value::String(s) => s.trim().trim_end_matches("px").parse::<f64>().ok()?,
        _ => return None,
    };
    if !f.is_finite() { return None; }
    Some(f.max(0.0).round().min(u32::MAX as f64) as u64)
}

fn coerce_bool(v: &Value) -> bool {
    match v {
        Value::Bool(b) => *b,
        Value::Number(n) => n.as_f64().is_some_and(|f| f != 0.0),
        Value::String(s) => matches!(s.trim().to_ascii_lowercase().as_str(), "true" | "1" | "yes"),
        _ => false,
    }
}

/// Make one run conform to the contract: strings stay, wrong types are
/// coerced or dropped (formatting only — the text is always kept).
fn repair_run(v: &mut Value) {
    if !v.is_object() {
        let t = match &*v { Value::String(s) => s.clone(), Value::Null => String::new(), other => other.to_string() };
        *v = json!({ "t": t });
        return;
    }
    let o = v.as_object_mut().expect("object");
    o.retain(|_, x| !x.is_null());
    let t = match o.remove("t").or_else(|| o.remove("text")) {
        Some(Value::String(s)) => s,
        Some(other) => other.to_string(),
        None => String::new(),
    };
    o.insert("t".into(), Value::String(t));
    for k in ["b", "i", "u", "s", "sup", "sub"] {
        if let Some(x) = o.get(k) { let b = coerce_bool(x); o.insert(k.into(), Value::Bool(b)); }
    }
    if let Some(x) = o.get("size").cloned() {
        match x {
            Value::Number(_) => {}
            Value::String(s) => match s.trim().trim_end_matches("pt").parse::<f64>() {
                Ok(f) if f.is_finite() => { o.insert("size".into(), json!(f)); }
                _ => { o.remove("size"); }
            },
            _ => { o.remove("size"); }
        }
    }
    for k in ["color", "font", "link", "hl"] {
        if o.get(k).is_some_and(|x| !x.is_string()) { o.remove(k); }
    }
}

fn repair_runs(o: &mut Map<String, Value>) {
    let runs = match o.remove("runs") {
        Some(Value::Array(a)) => a,
        Some(Value::Null) | None => match o.get("text") {
            Some(Value::String(s)) => vec![json!({ "t": s })],
            _ => Vec::new(),
        },
        Some(other) => vec![other],
    };
    let mut runs = runs;
    for r in &mut runs { repair_run(r); }
    o.insert("runs".into(), Value::Array(runs));
}

/// Repair a list of blocks in place; returns Err(description) for a block
/// that cannot be represented (the save is refused rather than losing it).
fn repair_blocks(blocks: &mut [Value], path: &str, notes: &mut Vec<String>) -> Result<(), String> {
    for (i, b) in blocks.iter_mut().enumerate() {
        let here = format!("{path}{}", i + 1);
        repair_block(b, &here, notes)?;
    }
    Ok(())
}

fn repair_block(v: &mut Value, here: &str, notes: &mut Vec<String>) -> Result<(), String> {
    if let Value::String(s) = v {
        *v = json!({ "type": "p", "runs": [{ "t": s.clone() }] });
        return Ok(());
    }
    let Some(o) = v.as_object_mut() else {
        return Err(format!("block {here} is not an object"));
    };
    o.retain(|_, x| !x.is_null());
    let ty = match o.get("type") {
        Some(Value::String(s)) => s.trim().to_ascii_lowercase(),
        _ if o.contains_key("rows") => "table".into(),
        _ if o.contains_key("src") => "img".into(),
        _ if o.contains_key("list") => "list".into(),
        _ => "p".into(),
    };
    let ty = match ty.as_str() {
        "paragraph" | "div" => "p".to_string(),
        "blockquote" => "quote".to_string(),
        "pre" => "code".to_string(),
        "image" => "img".to_string(),
        "page-break" | "pagebreak" | "break" => "pagebreak".to_string(),
        _ => ty,
    };
    match ty.as_str() {
        t if BLOCK_TYPES_WITH_RUNS.contains(&t) => {
            repair_runs(o);
            if o.get("align").is_some_and(|a| !a.is_string()) { o.remove("align"); }
        }
        "list" => {
            let kind = match o.get("list").and_then(Value::as_str).map(|s| s.to_ascii_lowercase()) {
                Some(s) if s == "ol" || s == "ordered" || s == "number" || s == "decimal" => "ol",
                _ => "ul",
            };
            o.insert("list".into(), json!(kind));
            let level = o.get("level").and_then(coerce_uint).unwrap_or(0).min(8);
            o.insert("level".into(), json!(level));
            repair_runs(o);
        }
        "table" => {
            let mut rows = match o.remove("rows") { Some(Value::Array(r)) => r, _ => Vec::new() };
            for (ri, row) in rows.iter_mut().enumerate() {
                if !row.is_array() { *row = Value::Array(vec![row.take()]); }
                let cells = row.as_array_mut().expect("array");
                for (ci, cell) in cells.iter_mut().enumerate() {
                    if !cell.is_object() || !cell.get("blocks").is_some_and(Value::is_array) {
                        // A bare value / runs-only cell becomes one paragraph.
                        let inner = match cell.take() {
                            Value::Object(mut co) => { co.remove("blocks"); json!([Value::Object(co)]) }
                            Value::Null => json!([]),
                            other => json!([other]),
                        };
                        *cell = json!({ "blocks": inner });
                    }
                    let co = cell.as_object_mut().expect("object");
                    co.retain(|_, x| !x.is_null());
                    match co.get("colspan").map(coerce_uint) {
                        Some(Some(n)) if n >= 1 => { co.insert("colspan".into(), json!(n)); }
                        Some(_) => { co.remove("colspan"); }
                        None => {}
                    }
                    let blocks = co.get_mut("blocks").and_then(Value::as_array_mut).expect("blocks");
                    repair_blocks(blocks, &format!("{here}, table cell {}:{} › ", ri + 1, ci + 1), notes)?;
                }
            }
            o.insert("rows".into(), Value::Array(rows));
        }
        "img" => {
            if !o.get("src").is_some_and(Value::is_string) {
                return Err(format!("block {here} is an image without a source"));
            }
            for k in ["w", "h"] {
                match o.get(k).map(coerce_uint) {
                    Some(Some(n)) => { o.insert(k.into(), json!(n)); }
                    Some(None) => { o.remove(k); }
                    None => {}
                }
            }
            match o.get("alt") {
                Some(Value::String(_)) | None => {}
                Some(other) => { let s = other.to_string(); o.insert("alt".into(), json!(s)); }
            }
        }
        "pagebreak" => {}
        "hr" => {
            // The model has no rule block; keep the position as an empty paragraph.
            notes.push(format!("Block {here}: a horizontal line was saved as an empty paragraph."));
            o.clear();
            o.insert("runs".into(), json!([]));
            o.insert("type".into(), json!("p"));
            return Ok(());
        }
        other => {
            // Unknown block type: keep its text as a paragraph.
            let has_text = o.get("runs").is_some_and(Value::is_array) || o.get("text").is_some_and(Value::is_string);
            if !has_text { return Err(format!("block {here} has unknown type {other:?} and no text")); }
            notes.push(format!("Block {here} of unknown type {other:?} was saved as a normal paragraph."));
            repair_runs(o);
            o.retain(|k, _| k == "runs" || k == "align");
            if o.get("align").is_some_and(|a| !a.is_string()) { o.remove("align"); }
            o.insert("type".into(), json!("p"));
            return Ok(());
        }
    }
    o.insert("type".into(), json!(ty));
    // Final check: the block must now deserialize into a typed variant.
    match serde_json::from_value::<Block>(v.clone()) {
        Ok(Block::Other(_)) | Err(_) => Err(format!("block {here} (type {ty:?}) has an invalid shape")),
        Ok(_) => Ok(()),
    }
}

/// Validate + normalise a word model JSON so that no block is silently
/// dropped by the writers (contract drift is repaired or rejected).
fn word_doc_from_value(mut model: Value, notes: &mut Vec<String>) -> Result<WordDoc, ApiError> {
    let Some(o) = model.as_object_mut() else {
        return Err(ApiError::unprocessable("Invalid document model: expected a JSON object."));
    };
    o.insert("kind".into(), json!("word"));
    if !o.get("meta").is_some_and(Value::is_object) { o.insert("meta".into(), json!({})); }
    let mut blocks = match o.remove("blocks") {
        Some(Value::Array(a)) => a,
        None | Some(Value::Null) => Vec::new(),
        Some(_) => return Err(ApiError::unprocessable("Invalid document model: \"blocks\" must be an array.")),
    };
    repair_blocks(&mut blocks, "#", notes)
        .map_err(|e| ApiError::unprocessable(format!("Invalid document model: {e}. Nothing was saved.")))?;
    o.insert("blocks".into(), Value::Array(blocks));
    serde_json::from_value(model).map_err(|e| ApiError::unprocessable(format!("Invalid document model: {e}")))
}

fn excel_book_from_value(model: Value) -> Result<ExcelBook, ApiError> {
    let mut book: ExcelBook = serde_json::from_value(model)
        .map_err(|e| ApiError::unprocessable(format!("Invalid spreadsheet model: {e}")))?;
    book.kind = "excel".into();
    Ok(book)
}

/// Plain-text rendering of a word model (mirrors wordModelToText in JS,
/// extended for tables, images and numbered lists).
fn word_to_text(doc: &WordDoc) -> String {
    fn runs_text(runs: &[Run]) -> String { runs.iter().map(|r| r.t.as_str()).collect() }
    fn blocks_text(blocks: &[Block], lines: &mut Vec<String>) {
        let mut counters = [0u32; 9];
        for b in blocks {
            if !matches!(b, Block::List { .. }) { counters = [0; 9]; }
            match b {
                Block::P { runs, .. } | Block::H1 { runs, .. } | Block::H2 { runs, .. }
                | Block::H3 { runs, .. } | Block::H4 { runs, .. } | Block::H5 { runs, .. }
                | Block::H6 { runs, .. } | Block::Quote { runs, .. } | Block::Code { runs, .. } => {
                    lines.push(runs_text(runs));
                }
                Block::List { list, level, runs } => {
                    let lvl = (*level as usize).min(8);
                    for c in counters.iter_mut().skip(lvl + 1) { *c = 0; }
                    counters[lvl] += 1;
                    let indent = "  ".repeat(lvl);
                    let marker = if list == "ol" { format!("{}.", counters[lvl]) } else { "•".into() };
                    lines.push(format!("{indent}{marker} {}", runs_text(runs)));
                }
                Block::Table { rows } => {
                    for row in rows {
                        let cells: Vec<String> = row.iter().map(|c| {
                            let mut inner = Vec::new();
                            blocks_text(&c.blocks, &mut inner);
                            inner.join(" ").replace(['\t', '\n'], " ")
                        }).collect();
                        lines.push(cells.join("\t"));
                    }
                }
                Block::Img { alt, .. } => {
                    lines.push(format!("[image{}]", alt.as_deref().filter(|a| !a.is_empty()).map(|a| format!(": {a}")).unwrap_or_default()));
                }
                Block::PageBreak => lines.push("\u{c}".into()),
                Block::Other(_) => {
                    // Loosely typed block: coerce it the same way the docx writer does.
                    if let Some(resolved) = crate::docx::resolve_block(b) {
                        if !matches!(resolved.as_ref(), Block::Other(_)) {
                            blocks_text(std::slice::from_ref(resolved.as_ref()), lines);
                        }
                    }
                }
            }
        }
    }
    let mut lines = Vec::new();
    blocks_text(&doc.blocks, &mut lines);
    let mut s = lines.join("\n");
    s.push('\n');
    s
}

/// Render a model (runs on a blocking thread). Returns bytes + notes.
fn render_model(kind: &str, format: &str, model: Value) -> Result<(Vec<u8>, Vec<String>), ApiError> {
    let mut notes = Vec::new();
    let write_failed = |e: String| ApiError::unprocessable(format!("Could not create the .{format} file: {e}"));
    let bytes = match kind {
        "word" => {
            let doc = word_doc_from_value(model, &mut notes)?;
            match format {
                "docx" => docx::write_docx(&doc).map_err(write_failed)?,
                "rtf" => textfmt::write_rtf(&doc),
                "txt" => word_to_text(&doc).into_bytes(),
                "md" => textfmt::write_markdown(&doc).into_bytes(),
                "html" => textfmt::write_html(&doc).into_bytes(),
                "tex" => latex::write_latex(&doc).into_bytes(),
                _ => unreachable!("validated by caller"),
            }
        }
        "excel" => {
            let mut book = excel_book_from_value(model)?;
            match format {
                "xlsx" => {
                    // Renamed sheets, truncated text, skipped merges... are
                    // reported to the client in X-Office-Warnings.
                    let (bytes, warns) = xlsx::write_xlsx_with_warnings(&book).map_err(write_failed)?;
                    notes.extend(warns);
                    bytes
                }
                "csv" | "tsv" => {
                    if book.sheets.is_empty() { return Ok((Vec::new(), notes)); }
                    let idx = (book.active as usize).min(book.sheets.len() - 1);
                    // CSV holds values only: export formula cells by their
                    // cached result, never the formula text.
                    for c in book.sheets[idx].cells.values_mut() { c.f = None; }
                    if book.sheets.len() > 1 {
                        notes.push(format!("Only the active sheet \"{}\" was saved; .{format} files hold a single sheet.", book.sheets[idx].name));
                    }
                    xlsx::sheet_to_delimited(&book, idx, if format == "tsv" { '\t' } else { ',' }).into_bytes()
                }
                _ => unreachable!("validated by caller"),
            }
        }
        "slides" => {
            let mut deck: SlidesDoc = serde_json::from_value(model)
                .map_err(|e| ApiError::unprocessable(format!("Invalid presentation model: {e}")))?;
            deck.meta.remove("warnings");
            match format {
                "pptx" => {
                    let (bytes, warns) = pptx::write_pptx(&deck).map_err(write_failed)?;
                    notes.extend(warns);
                    bytes
                }
                "md" => pptx::slides_to_markdown(&deck).into_bytes(),
                "txt" => pptx::slides_to_text(&deck).into_bytes(),
                _ => unreachable!("validated by caller"),
            }
        }
        "latex" => {
            let doc: LatexDoc = serde_json::from_value(model)
                .map_err(|e| ApiError::unprocessable(format!("Invalid LaTeX model: {e}")))?;
            match format {
                "tex" => doc.source.into_bytes(),
                "pdf" => latex::compile_pdf(&doc.source).map_err(|e| ApiError::unprocessable(format!("LaTeX compilation failed: {e}")))?,
                _ => {
                    let mut wd = latex::parse_latex(&doc.source);
                    if let Some(Value::Array(ws)) = wd.meta.remove("warnings") {
                        for w in ws { if let Value::String(s) = w { push_unique(&mut notes, s); } }
                    }
                    match format {
                        "docx" => docx::write_docx(&wd).map_err(write_failed)?,
                        "html" => textfmt::write_html(&wd).into_bytes(),
                        "md" => textfmt::write_markdown(&wd).into_bytes(),
                        "txt" => word_to_text(&wd).into_bytes(),
                        _ => unreachable!("validated by caller"),
                    }
                }
            }
        }
        _ => unreachable!("validated by caller"),
    };
    Ok((bytes, notes))
}

async fn api_render(body: Bytes) -> Result<Response, ApiError> {
    if body.is_empty() {
        return Err(ApiError::bad_request("Missing request body: expected JSON {kind, model, format, name}."));
    }
    let req: RenderReq = serde_json::from_slice(&body).map_err(|e| {
        if e.is_syntax() || e.is_eof() {
            ApiError::bad_request(format!("Request body is not valid JSON: {e}"))
        } else {
            ApiError::unprocessable(format!("Invalid render request: {e}"))
        }
    })?;
    drop(body);
    let kind = req.kind.trim().to_ascii_lowercase();
    let format = req.format.trim().trim_start_matches('.').to_ascii_lowercase();
    let latex_formats = render_latex_formats();
    let supported: &[&str] = match kind.as_str() {
        "word" => RENDER_WORD,
        "excel" => RENDER_EXCEL,
        "slides" => RENDER_SLIDES,
        "latex" => &latex_formats,
        _ => return Err(ApiError::bad_request(format!("Unknown document kind {:?} (expected \"word\", \"excel\", \"slides\" or \"latex\").", req.kind))),
    };
    if kind == "latex" && format == "pdf" && !supported.contains(&"pdf") {
        return Err(ApiError::unprocessable("PDF output needs a LaTeX engine (tectonic, pdflatex, xelatex or lualatex) installed on this computer and on PATH."));
    }
    let Some(&out_format) = supported.iter().find(|f| **f == format) else {
        return Err(ApiError::bad_request(format!(
            "Cannot save a {kind} document as .{format}. Supported formats: {}.", supported.join(", "))));
    };
    let model = req.model;
    let (bytes, notes) = run_blocking(move || render_model(&kind, out_format, model)).await?;

    let filename = output_filename(req.name.as_deref(), out_format);
    let size = bytes.len();
    let mut resp = Response::new(Body::from(bytes));
    let h = resp.headers_mut();
    h.insert(header::CONTENT_TYPE, HeaderValue::from_static(content_type_of(out_format)));
    h.insert(header::CONTENT_DISPOSITION, content_disposition(&filename));
    h.insert(X_OUTPUT_FORMAT, HeaderValue::from_static(out_format));
    h.insert(header::CACHE_CONTROL, HeaderValue::from_static("no-store"));
    if !notes.is_empty() {
        let enc = uri_component_encode(&Value::from(notes.clone()).to_string());
        if let Ok(v) = HeaderValue::from_str(&enc) { h.insert(X_OFFICE_WARNINGS, v); }
    }
    resp.extensions_mut().insert(LogNote(format!("{out_format} {}", human_size(size))));
    Ok(resp)
}

// =====================================================================
// Tests
// =====================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::{parse_docx, write_docx};
    use crate::model::{Block, ExcelBook, Run, WordDoc};
    use crate::textfmt::{csv_book, parse_rtf, parse_text};
    use crate::xlsx::{parse_delimited, parse_scalar, sheet_to_delimited, write_xlsx};
    use axum::http::Request as HttpRequest;
    use std::io::Write;
    use tower::ServiceExt;

    fn sample_word() -> WordDoc {
        WordDoc {
            kind: "word".into(),
            meta: Default::default(),
            blocks: vec![
                Block::H1 { runs: vec![Run::text("Title")], align: None },
                Block::P { runs: vec![Run { t: "bold".into(), b: true, ..Default::default() }, Run::text(" plain")], align: Some("center".into()) },
                Block::List { list: "ul".into(), level: 0, runs: vec![Run::text("item")] },
                Block::Table { rows: vec![vec![crate::model::TableCell { blocks: vec![Block::P { runs: vec![Run::text("c")], align: None }], colspan: None }]] },
                Block::PageBreak,
            ],
        }
    }

    // ---------- module-level tests (kept from the original suite) ----------

    #[test]
    fn csv_round_trip() {
        let text = "name,n\nA,1\nB,2\n";
        let book = parse_delimited(text, ',', "T");
        assert_eq!(book.sheets[0].cells["A2"].v.as_str().unwrap(), "A");
        assert_eq!(book.sheets[0].cells["B3"].v.as_f64().unwrap(), 2.0);
        let out = sheet_to_delimited(&book, 0, ',');
        assert!(out.contains("name,n"));
        assert!(out.contains("B,2"));
    }

    #[test]
    fn parse_scalar_kinds() {
        assert_eq!(parse_scalar("42").t, "n");
        assert_eq!(parse_scalar("=SUM(A1:A2)").f.as_deref(), Some("=SUM(A1:A2)"));
        assert_eq!(parse_scalar("TRUE").t, "b");
        assert_eq!(parse_scalar("hello").t, "s");
        assert_eq!(parse_scalar("2024-01-02").t, "d");
    }

    #[test]
    fn docx_round_trip() {
        let doc = sample_word();
        let bytes = write_docx(&doc).expect("write");
        assert_eq!(&bytes[..2], b"PK");
        let back = parse_docx(&bytes).expect("parse");
        let text: String = back.blocks.iter().flat_map(|b| b.para_runs().iter().map(|r| r.t.as_str())).collect();
        assert!(text.contains("Title"));
        assert!(text.contains("bold"));
        assert!(text.contains("item"));
        assert!(matches!(back.blocks[0], Block::H1 { .. }));
        assert!(matches!(back.blocks[4], Block::PageBreak));
        match &back.blocks[3] {
            Block::Table { rows } => {
                let cell_text: String = rows[0][0].blocks.iter()
                    .flat_map(|b| b.para_runs().iter().map(|r| r.t.as_str()))
                    .collect();
                assert_eq!(cell_text, "c");
            }
            other => panic!("expected table, got {other:?}"),
        }
    }

    #[test]
    fn xlsx_round_trip() {
        let book = ExcelBook {
            kind: "excel".into(),
            active: 0,
            warnings: vec![],
            sheets: vec![crate::model::Sheet {
                name: "S".into(),
                hidden: false,
                cells: [("A1".to_string(), crate::model::Cell::new(serde_json::json!(7), "n"))].into_iter().collect(),
                col_widths: None, merges: vec![], freeze: None, auto_filter: None, charts: vec![],
            }],
        };
        let bytes = write_xlsx(&book).expect("write");
        let back = crate::xlsx::parse_book(&bytes, None).expect("parse");
        let cell = &back.sheets[0].cells["A1"];
        assert_eq!(cell.v.as_f64().unwrap(), 7.0);
    }

    #[test]
    fn text_and_rtf_parse() {
        let doc = parse_text(b"hello\n\nworld", false);
        assert_eq!(doc.blocks.len(), 3);
        let rtf = br#"{\rtf1\ansi Hello \b bold\b0  world}"#;
        let doc = parse_rtf(rtf).expect("rtf");
        let text: String = doc.blocks.iter().flat_map(|b| b.para_runs().iter().map(|r| r.t.as_str())).collect();
        assert!(text.contains("Hello"));
        assert!(text.contains("bold"));
        let book = csv_book(b"a,b\n1,2\n", ',', "csv");
        assert_eq!(book.sheets[0].name, "csv");
    }

    #[test]
    fn format_detection() {
        assert_eq!(detect::format_of_name("a.DOCX"), Some("docx"));
        assert_eq!(detect::format_of_name("a.xls"), Some("xls"));
        assert_eq!(detect::resolve_format("x.bin", b"{\\rtf1 hi"), Some("rtf".to_string()));
        // Content wins over the extension: RTF bytes named .doc are RTF.
        assert_eq!(detect::resolve_format("x.doc", b"{\\rtf1 hi"), Some("rtf".to_string()));
        assert_eq!(detect::resolve_format("x.bin", &[0u8, 1, 2]), None);
        assert_eq!(detect::resolve_format("x.csv", b"a,b"), Some("csv".to_string()));
        // A real .docx renamed .doc is parsed as docx.
        let bytes = write_docx(&sample_word()).unwrap();
        assert_eq!(detect::resolve_format("renamed.doc", &bytes), Some("docx".to_string()));
    }

    // ---------- pure helpers ----------

    #[test]
    fn cli_parsing() {
        let s = |v: &[&str]| v.iter().map(|x| x.to_string()).collect::<Vec<_>>();
        let lo = |p: u16| SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), p);
        assert_eq!(parse_cli(&s(&[]), None), Ok(Cli::Run(lo(47831))));
        assert_eq!(parse_cli(&s(&["--port", "5000"]), None), Ok(Cli::Run(lo(5000))));
        assert_eq!(parse_cli(&s(&["--port=5001"]), None), Ok(Cli::Run(lo(5001))));
        assert_eq!(parse_cli(&s(&["-p", "5002"]), Some("127.0.0.1:9")), Ok(Cli::Run(lo(5002))));
        assert_eq!(parse_cli(&s(&[]), Some("5003")), Ok(Cli::Run(lo(5003))));
        assert_eq!(parse_cli(&s(&[]), Some("localhost:5004")), Ok(Cli::Run(lo(5004))));
        assert_eq!(parse_cli(&s(&["--addr", "[::1]:5005"]), None), Ok(Cli::Run("[::1]:5005".parse().unwrap())));
        assert_eq!(parse_cli(&s(&["--help"]), None), Ok(Cli::Help));
        assert_eq!(parse_cli(&s(&["-V"]), None), Ok(Cli::Version));
        assert!(parse_cli(&s(&["--port", "abc"]), None).is_err());
        assert!(parse_cli(&s(&["--port"]), None).is_err());
        assert!(parse_cli(&s(&["--bogus"]), None).is_err());
        assert!(parse_cli(&s(&[]), Some("nonsense:host:x")).is_err());
        let msg = friendly_bind_error(lo(47831), &std::io::Error::from(std::io::ErrorKind::AddrInUse));
        assert!(msg.contains("already in use") && msg.contains("--port 47832"), "{msg}");
    }

    #[test]
    fn file_name_decoding() {
        assert_eq!(decode_file_name(b"R%C3%A9sum%C3%A9%202026.docx"), "Résumé 2026.docx");
        assert_eq!(decode_file_name("报告.docx".as_bytes()), "报告.docx"); // raw UTF-8
        assert_eq!(decode_file_name(b"M\xe4rz.docx"), "März.docx"); // raw Latin-1 header bytes
        assert_eq!(decode_file_name(b"M%E4rz.docx"), "März.docx"); // %-encoded Latin-1
        assert_eq!(decode_file_name(b"100%.docx"), "100%.docx"); // invalid escape kept
        assert_eq!(decode_file_name(b"a%zzb%4"), "a%zzb%4");
        assert_eq!(decode_file_name(b"a%0Ab.txt"), "ab.txt"); // control chars dropped
    }

    #[test]
    fn host_and_origin_rules() {
        for ok in ["127.0.0.1", "127.0.0.1:47831", "localhost", "LOCALHOST:1", "[::1]", "[::1]:47831"] {
            assert!(host_allowed(ok.as_bytes()), "{ok}");
        }
        for bad in ["evil.example", "evil.example:47831", "127.0.0.1.evil.example", "localhost.evil:80", "127.0.0.2:47831",
                    "[::2]:1", "localhost:abc", "localhost:", "[::1", "0.0.0.0:47831", ""] {
            assert!(!host_allowed(bad.as_bytes()), "{bad}");
        }
        for ok in ["chrome-extension://omagnaefngeklmcndepoihhodddomkgo", "extension://abc", "moz-extension://1234-abcd"] {
            assert!(origin_allowed(ok.as_bytes()), "{ok}");
        }
        for bad in ["https://evil.example", "null", "chrome-extension://", "http://127.0.0.1:47831", "file://", "chrome-extension://a/b"] {
            assert!(!origin_allowed(bad.as_bytes()), "{bad}");
        }
    }

    #[test]
    fn filenames_and_disposition() {
        assert_eq!(output_filename(Some("Résumé (final).docx"), "docx"), "Résumé (final).docx");
        assert_eq!(output_filename(Some("report.rtf"), "docx"), "report.docx");
        assert_eq!(output_filename(Some("v1.2 notes"), "md"), "v1.2 notes.md");
        assert_eq!(output_filename(Some("C:\\x\\a:b*c?.xlsx"), "csv"), "a_b_c_.csv");
        assert_eq!(output_filename(None, "xlsx"), "document.xlsx");
        assert_eq!(output_filename(Some("  ..  "), "txt"), "document.txt");
        assert_eq!(ascii_fallback("Résumé (final).docx"), "Resume (final).docx");
        assert_eq!(ascii_fallback("报告 \"x\";%.docx"), "__ _x___.docx");
        assert_eq!(rfc5987_encode("Отчёт 1.docx"), "%D0%9E%D1%82%D1%87%D1%91%D1%82%201.docx");
        let cd = content_disposition("Отчёт за квартал.docx");
        let s = cd.to_str().unwrap();
        assert!(s.starts_with("attachment; filename=\""));
        assert!(s.contains("filename*=UTF-8''%D0%9E"));
    }

    #[test]
    fn word_model_repair() {
        let mut notes = Vec::new();
        let model = json!({ "kind": "word", "blocks": [
            { "type": "list", "list": "ol", "level": -1, "runs": [{ "t": "negative level item" }] },
            { "type": "img", "src": "data:image/png;base64,AAAA", "w": 100.5, "h": null },
            { "type": "hr" },
            { "type": "callout", "runs": [{ "t": "kept text", "b": 1 }] },
            { "type": "p", "runs": [{ "t": 5, "size": "12pt", "color": null }], "align": null },
            "bare string",
            { "type": "table", "rows": [[ { "blocks": [ { "type": "weird", "text": "cell text" } ], "colspan": 2.0 }, "plain cell" ]] },
        ]});
        let doc = word_doc_from_value(model, &mut notes).expect("repairable");
        assert_eq!(doc.blocks.len(), 7);
        assert!(doc.blocks.iter().all(|b| !matches!(b, Block::Other(_))));
        assert!(matches!(&doc.blocks[0], Block::List { level: 0, list, .. } if list == "ol"));
        assert!(matches!(&doc.blocks[1], Block::Img { w: Some(101), h: None, .. }) || matches!(&doc.blocks[1], Block::Img { w: Some(100), .. }));
        let txt = word_to_text(&doc);
        for needle in ["negative level item", "kept text", "5", "bare string", "cell text", "plain cell", "[image]"] {
            assert!(txt.contains(needle), "{needle} missing from {txt:?}");
        }
        assert!(notes.iter().any(|n| n.contains("callout")));

        let bad = json!({ "blocks": [ { "type": "img" } ] });
        let e = word_doc_from_value(bad, &mut notes).unwrap_err();
        assert_eq!(e.status, StatusCode::UNPROCESSABLE_ENTITY);
        assert!(e.msg.contains("#1"), "{}", e.msg);
    }

    #[test]
    fn word_text_numbering() {
        let doc = WordDoc { kind: "word".into(), meta: Default::default(), blocks: vec![
            Block::List { list: "ol".into(), level: 0, runs: vec![Run::text("a")] },
            Block::List { list: "ol".into(), level: 1, runs: vec![Run::text("a.i")] },
            Block::List { list: "ol".into(), level: 0, runs: vec![Run::text("b")] },
        ]};
        assert_eq!(word_to_text(&doc), "1. a\n  1. a.i\n2. b\n");
    }

    // ---------- HTTP-level tests (full middleware stack) ----------

    struct Resp { status: StatusCode, headers: HeaderMap, body: Vec<u8> }
    impl Resp {
        fn json(&self) -> Value { serde_json::from_slice(&self.body).unwrap_or_else(|e| panic!("not JSON ({e}): {:?}", String::from_utf8_lossy(&self.body))) }
    }

    async fn call(req: HttpRequest<Body>) -> Resp {
        let r = app().oneshot(req).await.expect("infallible");
        let status = r.status();
        let headers = r.headers().clone();
        let body = axum::body::to_bytes(r.into_body(), usize::MAX).await.unwrap().to_vec();
        Resp { status, headers, body }
    }

    fn parse_req(name: &str, body: Vec<u8>) -> HttpRequest<Body> {
        HttpRequest::post("/api/parse").header("host", "127.0.0.1:47831")
            .header("content-type", "application/octet-stream")
            .header("x-file-name", HeaderValue::from_bytes(name.as_bytes()).unwrap())
            .body(Body::from(body)).unwrap()
    }

    fn render_req(v: Value) -> HttpRequest<Body> {
        HttpRequest::post("/api/render").header("content-type", "application/json")
            .body(Body::from(v.to_string())).unwrap()
    }

    fn assert_json_error(r: &Resp, status: StatusCode) -> String {
        assert_eq!(r.status, status, "body: {}", String::from_utf8_lossy(&r.body));
        assert!(r.headers[header::CONTENT_TYPE].to_str().unwrap().starts_with("application/json"));
        r.json()["error"].as_str().expect("error string").to_string()
    }

    #[tokio::test]
    async fn health_shape() {
        let r = call(HttpRequest::get("/health").body(Body::empty()).unwrap()).await;
        assert_eq!(r.status, StatusCode::OK);
        let j = r.json();
        assert_eq!(j["ok"], true);
        assert_eq!(j["name"], "office-backend");
        assert_eq!(j["api"], 2);
        assert_eq!(j["version"], VERSION);
        assert_eq!(j["render"]["word"], json!(["docx", "rtf", "txt", "md", "html", "tex"]));
        assert_eq!(j["render"]["slides"], json!(["pptx", "md", "txt"]));
        assert!(j["render"]["latex"].as_array().unwrap().iter().any(|x| x == "docx"));
        assert_eq!(j["render"]["excel"], json!(["xlsx", "csv", "tsv"]));
        for f in ["docx", "doc", "rtf", "odt", "xlsx", "xls", "csv", "tsv", "pptx", "tex"] {
            assert!(j["parse"].as_array().unwrap().iter().any(|x| x == f), "{f}");
        }
    }

    #[tokio::test]
    async fn security_origin_and_host() {
        // Web origin: refused before any parser runs (also simple text/plain POSTs).
        let mut req = parse_req("a.txt", b"hello".to_vec());
        req.headers_mut().insert(header::ORIGIN, HeaderValue::from_static("https://evil.example"));
        req.headers_mut().insert(header::CONTENT_TYPE, HeaderValue::from_static("text/plain"));
        let r = call(req).await;
        assert_json_error(&r, StatusCode::FORBIDDEN);
        assert!(r.headers.get(header::ACCESS_CONTROL_ALLOW_ORIGIN).is_none());

        // Preflight from a web origin: 403, no CORS grant.
        let pre = HttpRequest::builder().method(Method::OPTIONS).uri("/api/parse")
            .header("origin", "https://evil.example").header("access-control-request-method", "POST")
            .header("access-control-request-headers", "x-file-name").body(Body::empty()).unwrap();
        let r = call(pre).await;
        assert_eq!(r.status, StatusCode::FORBIDDEN);
        assert!(r.headers.get(header::ACCESS_CONTROL_ALLOW_ORIGIN).is_none());

        // DNS rebinding: foreign Host header.
        let r = call(HttpRequest::get("/health").header("host", "attacker.example:47831").body(Body::empty()).unwrap()).await;
        assert_json_error(&r, StatusCode::FORBIDDEN);

        // Extension origin: allowed and CORS reflected with the needed headers.
        let ext = "chrome-extension://omagnaefngeklmcndepoihhodddomkgo";
        let pre = HttpRequest::builder().method(Method::OPTIONS).uri("/api/render")
            .header("host", "localhost:47831").header("origin", ext)
            .header("access-control-request-method", "POST")
            .header("access-control-request-headers", "content-type,x-file-name").body(Body::empty()).unwrap();
        let r = call(pre).await;
        assert!(r.status.is_success(), "{}", r.status);
        assert_eq!(r.headers[header::ACCESS_CONTROL_ALLOW_ORIGIN], ext);
        let allow = r.headers[header::ACCESS_CONTROL_ALLOW_HEADERS].to_str().unwrap().to_ascii_lowercase();
        assert!(allow.contains("x-file-name") && allow.contains("content-type"), "{allow}");

        let mut req = parse_req("a.txt", b"hello".to_vec());
        req.headers_mut().insert(header::ORIGIN, HeaderValue::from_static("chrome-extension://omagnaefngeklmcndepoihhodddomkgo"));
        let r = call(req).await;
        assert_eq!(r.status, StatusCode::OK);
        assert_eq!(r.headers[header::ACCESS_CONTROL_ALLOW_ORIGIN], ext);
        let expose = r.headers[header::ACCESS_CONTROL_EXPOSE_HEADERS].to_str().unwrap().to_ascii_lowercase();
        assert!(expose.contains("content-disposition") && expose.contains("x-output-format"), "{expose}");

        // Web-page subresource probe without Origin (e.g. <img src=.../health>).
        let r = call(HttpRequest::get("/health").header("sec-fetch-site", "cross-site").header("sec-fetch-dest", "image")
            .body(Body::empty()).unwrap()).await;
        assert_json_error(&r, StatusCode::FORBIDDEN);
        // ...but the user opening /health in a tab is fine.
        let r = call(HttpRequest::get("/health").header("sec-fetch-site", "none").body(Body::empty()).unwrap()).await;
        assert_eq!(r.status, StatusCode::OK);

        // Edge / Firefox extension origins are fine too; no Origin (curl) is fine.
        for o in ["extension://abcdef", "moz-extension://1b2c-3d"] {
            let r = call(HttpRequest::get("/health").header("origin", o).body(Body::empty()).unwrap()).await;
            assert_eq!(r.status, StatusCode::OK, "{o}");
        }
    }

    #[tokio::test]
    async fn parse_non_ascii_names() {
        // Percent-encoded (what the extension sends), raw Latin-1 and raw UTF-8.
        let bytes = write_docx(&sample_word()).unwrap();
        for name in [&b"R%C3%A9sum%C3%A9%20%D0%BE%D1%82%D1%87%D1%91%D1%82.docx"[..], b"R\xe9sum\xe9.docx", "报告.docx".as_bytes()] {
            let req = HttpRequest::post("/api/parse").header("x-file-name", HeaderValue::from_bytes(name).unwrap())
                .body(Body::from(bytes.clone())).unwrap();
            let r = call(req).await;
            assert_eq!(r.status, StatusCode::OK, "{:?}: {}", String::from_utf8_lossy(name), String::from_utf8_lossy(&r.body));
            let j = r.json();
            assert_eq!(j["kind"], "word");
            assert_eq!(j["format"], "docx");
            assert!(j["warnings"].is_array());
        }
        // No name at all: the content still decides.
        let r = call(HttpRequest::post("/api/parse").body(Body::from(bytes.clone())).unwrap()).await;
        assert_eq!(r.json()["format"], "docx");
        let r = call(parse_req("R%C3%A9sum%C3%A9.docx", bytes)).await;
        assert_eq!(r.json()["name"], "Résumé.docx");
    }

    #[tokio::test]
    async fn parse_routes_and_warnings() {
        // .docx renamed .doc -> docx with a note.
        let bytes = write_docx(&sample_word()).unwrap();
        let j = call(parse_req("renamed.doc", bytes)).await.json();
        assert_eq!(j["format"], "docx");
        assert!(j["warnings"][0].as_str().unwrap().contains(".doc"));

        // CSV with semicolons is sniffed; TSV uses tabs; sheet named after the file.
        let j = call(parse_req("Umsatz%20Q3.csv", b"a;b\n1;2\n".to_vec())).await.json();
        assert_eq!(j["kind"], "excel");
        assert_eq!(j["model"]["sheets"][0]["name"], "Umsatz Q3");
        assert_eq!(j["model"]["sheets"][0]["cells"]["B2"]["v"], 2.0);
        let j = call(parse_req("t.tsv", b"x\ty,z\n".to_vec())).await.json();
        assert_eq!(j["model"]["sheets"][0]["cells"]["B1"]["v"], "y,z");
        // Windows-1252 CSV: encoding note lifted out of the model.
        let j = call(parse_req("w.csv", b"name\nM\xfcller\n".to_vec())).await.json();
        assert!(j["model"].get("warnings").is_none());
        assert_eq!(j["model"]["sheets"][0]["cells"]["A2"]["v"], "Müller");

        // CSV saved as .xls -> opened as csv with a note.
        let j = call(parse_req("export.xls", b"a,b\n1,2\n".to_vec())).await.json();
        assert_eq!(j["format"], "csv");
        assert!(!j["warnings"].as_array().unwrap().is_empty());

        // HTML saved as .xls -> clear 415.
        let r = call(parse_req("export.xls", b"<html><table><tr><td>1</td></tr></table></html>".to_vec())).await;
        assert!(assert_json_error(&r, StatusCode::UNSUPPORTED_MEDIA_TYPE).contains("HTML"));

        // Legacy .doc -> lossy warning, warnings removed from meta.
        let ole = crate::detect::fake_ole("WordDocument");
        let j = call(parse_req("old.doc", ole)).await.json();
        assert_eq!(j["format"], "doc");
        assert!(j["warnings"].as_array().unwrap().iter().any(|w| w.as_str().unwrap().contains(".doc")));
        assert!(j["model"]["meta"].get("warnings").is_none());

        // Unknown type / empty body.
        let r = call(parse_req("x.pdf", b"%PDF-1.7".to_vec())).await;
        assert_json_error(&r, StatusCode::UNSUPPORTED_MEDIA_TYPE);
        let r = call(parse_req("x.docx", Vec::new())).await;
        assert_json_error(&r, StatusCode::BAD_REQUEST);
        // Corrupt docx.
        let r = call(parse_req("x.docx", b"PK\x03\x04garbage".to_vec())).await;
        assert_json_error(&r, StatusCode::UNPROCESSABLE_ENTITY);
    }

    #[tokio::test]
    async fn parse_zip_bomb_rejected() {
        let mut w = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let opts = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        w.start_file("[Content_Types].xml", opts).unwrap();
        w.write_all(b"wordprocessingml.document.main+xml").unwrap();
        w.start_file("word/document.xml", opts).unwrap();
        let chunk = vec![b' '; 1 << 20];
        for _ in 0..40 { w.write_all(&chunk).unwrap(); }
        let bomb = w.finish().unwrap().into_inner();
        let r = call(parse_req("bomb.docx", bomb)).await;
        assert!(assert_json_error(&r, StatusCode::PAYLOAD_TOO_LARGE).contains("zip bomb"));
    }

    #[tokio::test]
    async fn errors_are_json() {
        // Body over the limit (Content-Length known up front).
        let big = vec![b'a'; MAX_BODY + 1];
        let r = call(parse_req("big.txt", big)).await;
        assert!(assert_json_error(&r, StatusCode::PAYLOAD_TOO_LARGE).contains("100 MB"));
        // Bad JSON, wrong shape, missing body, unknown route, wrong method.
        let r = call(HttpRequest::post("/api/render").body(Body::from("{nope")).unwrap()).await;
        assert_json_error(&r, StatusCode::BAD_REQUEST);
        let r = call(render_req(json!({ "kind": "word", "model": {} }))).await;
        assert!(assert_json_error(&r, StatusCode::UNPROCESSABLE_ENTITY).contains("format"));
        let r = call(HttpRequest::post("/api/render").body(Body::empty()).unwrap()).await;
        assert_json_error(&r, StatusCode::BAD_REQUEST);
        let r = call(HttpRequest::get("/nope").body(Body::empty()).unwrap()).await;
        assert_json_error(&r, StatusCode::NOT_FOUND);
        let r = call(HttpRequest::get("/api/parse").body(Body::empty()).unwrap()).await;
        assert_json_error(&r, StatusCode::METHOD_NOT_ALLOWED);
        // Unsupported output format names the supported ones.
        let r = call(render_req(json!({ "kind": "excel", "model": { "sheets": [] }, "format": "ods" }))).await;
        assert!(assert_json_error(&r, StatusCode::BAD_REQUEST).contains("xlsx, csv, tsv"));
        // Irreparable block -> 422, not a silent empty paragraph.
        let r = call(render_req(json!({ "kind": "word", "model": { "blocks": [ { "type": "img" } ] }, "format": "docx" }))).await;
        assert_json_error(&r, StatusCode::UNPROCESSABLE_ENTITY);
    }

    #[tokio::test]
    async fn panics_become_json_500() {
        async fn boom() -> &'static str { panic!("kaboom") }
        let app = Router::new().route("/boom", get(boom))
            .layer(CatchPanicLayer::custom(panic_response))
            .layer(middleware::from_fn(json_errors));
        let r = app.oneshot(HttpRequest::get("/boom").body(Body::empty()).unwrap()).await.unwrap();
        assert_eq!(r.status(), StatusCode::INTERNAL_SERVER_ERROR);
        let body = axum::body::to_bytes(r.into_body(), usize::MAX).await.unwrap();
        let j: Value = serde_json::from_slice(&body).unwrap();
        assert!(j["error"].as_str().unwrap().contains("kaboom"));
        // Panics inside blocking jobs too.
        let e = run_blocking::<(), _>(|| panic!("blocking kaboom")).await.unwrap_err();
        assert_eq!(e.status, StatusCode::INTERNAL_SERVER_ERROR);
        assert!(e.msg.contains("blocking kaboom"));
    }

    #[tokio::test]
    async fn render_all_formats() {
        let model = serde_json::to_value(sample_word()).unwrap();
        for (fmt, ct) in [("docx", "wordprocessingml"), ("rtf", "application/rtf"), ("txt", "text/plain"),
                          ("md", "text/markdown"), ("html", "text/html")] {
            let r = call(render_req(json!({ "kind": "word", "model": model, "format": fmt, "name": "Отчёт за квартал.rtf" }))).await;
            assert_eq!(r.status, StatusCode::OK, "{fmt}: {}", String::from_utf8_lossy(&r.body));
            assert!(r.headers[header::CONTENT_TYPE].to_str().unwrap().contains(ct), "{fmt}");
            assert_eq!(r.headers[X_OUTPUT_FORMAT], fmt);
            let cd = r.headers[header::CONTENT_DISPOSITION].to_str().unwrap();
            assert!(cd.contains(&format!("filename*=UTF-8''%D0%9E%D1%82%D1%87%D1%91%D1%82%20%D0%B7%D0%B0%20%D0%BA%D0%B2%D0%B0%D1%80%D1%82%D0%B0%D0%BB.{fmt}")), "{cd}");
            assert!(!r.body.is_empty());
            let text = String::from_utf8_lossy(&r.body);
            if fmt != "docx" { assert!(text.contains("Title") && text.contains("item"), "{fmt}: {text}"); }
        }
        // Excel: csv exports cached values (never formula text) of the active sheet.
        let book = json!({ "kind": "excel", "active": 1, "sheets": [
            { "name": "A", "cells": { "A1": { "v": "first", "t": "s" } } },
            { "name": "B", "cells": { "A1": { "v": 2, "t": "n" }, "B1": { "v": 4, "t": "n", "f": "=A1*2" } } } ] });
        let r = call(render_req(json!({ "kind": "excel", "model": book, "format": "csv", "name": "Bericht März.xlsx" }))).await;
        assert_eq!(r.status, StatusCode::OK);
        assert_eq!(String::from_utf8_lossy(&r.body).trim(), "2,4");
        let cd = r.headers[header::CONTENT_DISPOSITION].to_str().unwrap();
        assert!(cd.contains("filename=\"Bericht Marz.csv\""), "{cd}");
        assert!(r.headers.contains_key(X_OFFICE_WARNINGS));
        let r = call(render_req(json!({ "kind": "excel", "model": book, "format": "xlsx", "name": "b" }))).await;
        assert_eq!(r.status, StatusCode::OK);
        assert_eq!(&r.body[..2], b"PK");
        let r = call(render_req(json!({ "kind": "excel", "model": book, "format": "tsv" }))).await;
        assert!(r.headers[header::CONTENT_TYPE].to_str().unwrap().starts_with("text/tab-separated-values"));
    }

    #[tokio::test]
    async fn server_on_ephemeral_port() {
        // Real TCP: a request with a foreign Host is refused, loopback works.
        let listener = tokio::net::TcpListener::bind("127.0.0.1:0").await.unwrap();
        let addr = listener.local_addr().unwrap();
        let server = tokio::spawn(async move { axum::serve(listener, app()).await.unwrap() });
        use tokio::io::{AsyncReadExt, AsyncWriteExt};
        async fn raw(addr: SocketAddr, req: String) -> String {
            let mut s = tokio::net::TcpStream::connect(addr).await.unwrap();
            s.write_all(req.as_bytes()).await.unwrap();
            let mut out = Vec::new();
            s.read_to_end(&mut out).await.unwrap();
            String::from_utf8_lossy(&out).into_owned()
        }
        let ok = raw(addr, format!("GET /health HTTP/1.1\r\nHost: 127.0.0.1:{}\r\nConnection: close\r\n\r\n", addr.port())).await;
        assert!(ok.starts_with("HTTP/1.1 200"), "{ok}");
        let bad = raw(addr, "GET /health HTTP/1.1\r\nHost: rebind.example\r\nConnection: close\r\n\r\n".into()).await;
        assert!(bad.starts_with("HTTP/1.1 403"), "{bad}");
        server.abort();
        // Binding the same port again gives a friendly error, not a panic.
        let again = std::net::TcpListener::bind(addr);
        let _ = again; // port is free again after abort on most OSes; only check the message path:
        let msg = friendly_bind_error(addr, &std::io::Error::from(std::io::ErrorKind::AddrInUse));
        assert!(msg.contains(&addr.port().to_string()));
    }
}
