//! Input safety limits (owned by the api workstream).
//!
//! `check_zip` is called by main.rs on every ZIP container (docx/xlsx/xlsm/
//! ods/odt, and anything else starting with "PK") BEFORE format sniffing or
//! parsing, so zip bombs are rejected before any parser inflates them.
//!
//! What is checked:
//! 1. The end-of-central-directory record (and its ZIP64 variant) is located
//!    by hand and the entry count is checked *before* the central directory
//!    is materialised (a 100 MB archive could otherwise declare millions of
//!    entries and make the directory itself a memory bomb).
//! 2. The declared uncompressed sizes are summed and checked.
//! 3. Every entry is actually inflated into a sink (nothing is kept in
//!    memory) with a hard byte budget, because headers can lie about sizes
//!    and the zip readers used by the parsers do not cap output at the
//!    declared size. This also yields the true compression ratio per entry.

use std::io::{self, Cursor, Read};

/// Max total uncompressed size of all entries in an uploaded ZIP container.
/// (xlsx.rs also uses this as its own per-package budget.)
pub const MAX_UNZIPPED_TOTAL: u64 = 512 * 1024 * 1024;
/// Max number of entries in an uploaded ZIP container.
pub const MAX_ZIP_ENTRIES: usize = 20_000;
/// Entries that inflate beyond this size must have a sane compression ratio.
pub const RATIO_CHECK_MIN_SIZE: u64 = 16 * 1024 * 1024;
/// Max inflated/compressed ratio for large entries. Real XML compresses
/// ~5-40x; deflate's theoretical maximum is ~1032x (classic bombs sit there).
pub const MAX_COMPRESSION_RATIO: u64 = 200;
/// Max HTTP request body (the upload itself).
pub const MAX_BODY: usize = 100 * 1024 * 1024;

/// Why a ZIP container was rejected.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ZipCheckError {
    /// Exceeds a safety limit (maps to HTTP 413).
    TooLarge(String),
    /// Not a readable ZIP archive (maps to HTTP 422).
    Corrupt(String),
}

impl ZipCheckError {
    pub fn message(&self) -> &str {
        match self { ZipCheckError::TooLarge(m) | ZipCheckError::Corrupt(m) => m }
    }
}

impl std::fmt::Display for ZipCheckError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result { f.write_str(self.message()) }
}

/// Summary of a container that passed the checks.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ZipStats {
    pub entries: usize,
    pub declared_total: u64,
    pub inflated_total: u64,
}

/// CONTRACT: validate a ZIP container against the limits above.
#[allow(dead_code)] // contract API; main.rs uses check_zip_detailed for 413-vs-422
pub fn check_zip(bytes: &[u8]) -> Result<(), String> {
    check_zip_detailed(bytes).map(|_| ()).map_err(|e| e.to_string())
}

fn mib(n: u64) -> u64 { n / (1024 * 1024) }

/// Like `check_zip`, but distinguishes "too large" from "corrupt" so the
/// HTTP layer can choose 413 vs 422, and returns statistics.
pub fn check_zip_detailed(bytes: &[u8]) -> Result<ZipStats, ZipCheckError> {
    check_zip_with(bytes, MAX_UNZIPPED_TOTAL, MAX_ZIP_ENTRIES)
}

/// Implementation with explicit limits (used by tests).
pub fn check_zip_with(bytes: &[u8], max_total: u64, max_entries: usize) -> Result<ZipStats, ZipCheckError> {
    let corrupt = |m: &str| ZipCheckError::Corrupt(format!("The file is not a valid ZIP-based document ({m})."));
    let declared_entries = eocd_entry_count(bytes).ok_or_else(|| corrupt("no end-of-central-directory record"))?;
    if declared_entries > max_entries as u64 {
        return Err(ZipCheckError::TooLarge(format!(
            "The file contains {declared_entries} archive entries (limit {max_entries}); refusing to open it.")));
    }

    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| corrupt(&e.to_string()))?;
    let n = archive.len();
    if n > max_entries {
        return Err(ZipCheckError::TooLarge(format!(
            "The file contains {n} archive entries (limit {max_entries}); refusing to open it.")));
    }

    // Pass 1: declared sizes (cheap, catches honest-but-huge archives).
    let mut declared_total: u64 = 0;
    for i in 0..n {
        let f = archive.by_index_raw(i).map_err(|e| corrupt(&e.to_string()))?;
        declared_total = declared_total.saturating_add(f.size());
    }
    if declared_total > max_total {
        return Err(ZipCheckError::TooLarge(format!(
            "The document would expand to {} MB when unpacked (limit {} MB); refusing to open it.",
            mib(declared_total), mib(max_total))));
    }

    // Pass 2: inflate every entry into a sink with a hard budget. Headers can
    // lie, so the real size is what counts.
    let mut inflated_total: u64 = 0;
    for i in 0..n {
        let mut f = match archive.by_index(i) {
            Ok(f) => f,
            // Encrypted entries / unsupported compression methods: the parsers
            // cannot inflate them either, so they cannot bomb; rely on the
            // declared size counted above.
            Err(_) => continue,
        };
        if f.is_dir() { continue; }
        let compressed = f.compressed_size();
        // Read at most one byte past the remaining budget.
        let budget = (max_total - inflated_total).saturating_add(1);
        let mut limited = (&mut f).take(budget);
        let copied = io::copy(&mut limited, &mut io::sink());
        let got = budget - limited.limit();
        let size_lie = copied.as_ref().err()
            .is_some_and(|e| e.to_string().to_ascii_lowercase().contains("declared"));
        if size_lie {
            // The zip reader stops with an error once an entry inflates past
            // its declared size: the headers lie, which a genuine document
            // never does. (Other read errors — bad CRC, truncated data — are
            // left for the parser to report.)
            return Err(ZipCheckError::TooLarge(format!(
                "Archive entry \"{}\" is larger than its declared size; refusing to open it (possible zip bomb).",
                f.name().map(|n| n.to_string()).unwrap_or_default())));
        }
        inflated_total = inflated_total.saturating_add(got);
        if inflated_total > max_total {
            return Err(ZipCheckError::TooLarge(format!(
                "The document expands to more than {} MB when unpacked; refusing to open it (possible zip bomb).",
                mib(max_total))));
        }
        if got > RATIO_CHECK_MIN_SIZE && got / compressed.max(1) > MAX_COMPRESSION_RATIO {
            return Err(ZipCheckError::TooLarge(format!(
                "Archive entry \"{}\" has an implausible compression ratio ({}:1); refusing to open it (possible zip bomb).",
                f.name().map(|n| n.to_string()).unwrap_or_default(), got / compressed.max(1))));
        }
    }
    Ok(ZipStats { entries: n, declared_total, inflated_total })
}

fn u16_at(b: &[u8], i: usize) -> Option<u16> { b.get(i..i + 2).map(|s| u16::from_le_bytes([s[0], s[1]])) }
fn u32_at(b: &[u8], i: usize) -> Option<u32> { b.get(i..i + 4).map(|s| u32::from_le_bytes([s[0], s[1], s[2], s[3]])) }
fn u64_at(b: &[u8], i: usize) -> Option<u64> {
    b.get(i..i + 8).map(|s| u64::from_le_bytes([s[0], s[1], s[2], s[3], s[4], s[5], s[6], s[7]]))
}

/// Find the end-of-central-directory record (searching backwards over the
/// max comment length) and return the total entry count it declares,
/// following the ZIP64 locator when the classic field is saturated.
pub fn eocd_entry_count(b: &[u8]) -> Option<u64> {
    const EOCD: [u8; 4] = [0x50, 0x4b, 0x05, 0x06];
    if b.len() < 22 { return None; }
    let min_start = b.len().saturating_sub(22 + 0xFFFF);
    let mut pos = b.len() - 22;
    let eocd = loop {
        if b[pos..pos + 4] == EOCD { break pos; }
        if pos == min_start { return None; }
        pos -= 1;
    };
    let total = u16_at(b, eocd + 10)? as u64;
    let cd_size = u32_at(b, eocd + 12)?;
    let cd_off = u32_at(b, eocd + 16)?;
    if total != 0xFFFF && cd_size != 0xFFFF_FFFF && cd_off != 0xFFFF_FFFF {
        return Some(total);
    }
    // ZIP64: locator sits immediately before the classic EOCD.
    let loc = eocd.checked_sub(20)?;
    if b.get(loc..loc + 4)? != [0x50, 0x4b, 0x06, 0x07] { return Some(total); }
    let z64 = usize::try_from(u64_at(b, loc + 8)?).ok()?;
    if b.get(z64..z64 + 4)? != [0x50, 0x4b, 0x06, 0x06] { return None; }
    u64_at(b, z64 + 32)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::write::SimpleFileOptions;

    fn make_zip(entries: &[(&str, Vec<u8>)]) -> Vec<u8> {
        let mut w = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        for (name, data) in entries {
            w.start_file(*name, opts).unwrap();
            w.write_all(data).unwrap();
        }
        w.finish().unwrap().into_inner()
    }

    #[test]
    fn accepts_normal_archive() {
        let z = make_zip(&[("a.xml", b"<a>hello</a>".to_vec()), ("b.xml", vec![b'x'; 1000])]);
        let s = check_zip_detailed(&z).expect("ok");
        assert_eq!(s.entries, 2);
        assert_eq!(s.inflated_total, 12 + 1000);
        assert!(check_zip(&z).is_ok());
    }

    #[test]
    fn rejects_garbage() {
        assert!(matches!(check_zip_detailed(b"PK\x03\x04 not really"), Err(ZipCheckError::Corrupt(_))));
        assert!(check_zip(b"").is_err());
    }

    #[test]
    fn rejects_too_many_entries() {
        let entries: Vec<(String, Vec<u8>)> = (0..30).map(|i| (format!("f{i}"), vec![1u8])).collect();
        let refs: Vec<(&str, Vec<u8>)> = entries.iter().map(|(n, d)| (n.as_str(), d.clone())).collect();
        let z = make_zip(&refs);
        assert_eq!(eocd_entry_count(&z), Some(30));
        assert!(matches!(check_zip_with(&z, 1 << 30, 10), Err(ZipCheckError::TooLarge(_))));
    }

    #[test]
    fn rejects_oversized_total() {
        let z = make_zip(&[("big.xml", vec![b'a'; 200_000])]);
        let e = check_zip_with(&z, 100_000, 100).unwrap_err();
        assert!(matches!(e, ZipCheckError::TooLarge(_)), "{e:?}");
    }

    #[test]
    fn rejects_high_ratio_bomb() {
        // 40 MB of zeros deflates to ~40 KB (ratio ~1000:1).
        let z = make_zip(&[("word/document.xml", vec![0u8; 40 * 1024 * 1024])]);
        assert!(z.len() < 1024 * 1024);
        let e = check_zip_detailed(&z).unwrap_err();
        assert!(e.message().contains("compression ratio"), "{e}");
    }

    #[test]
    fn detects_lying_size_headers() {
        // Patch the declared uncompressed size (central directory + local
        // header) down to 1 byte: the real inflate must still be counted.
        let mut z = make_zip(&[("x.xml", vec![b'q'; 300_000])]);
        for sig in [[0x50u8, 0x4b, 0x01, 0x02], [0x50, 0x4b, 0x03, 0x04]] {
            let off = if sig[2] == 1 { 24 } else { 22 };
            let pos = z.windows(4).position(|w| w == sig).unwrap();
            z[pos + off..pos + off + 4].copy_from_slice(&1u32.to_le_bytes());
        }
        let r = check_zip_with(&z, 100_000, 100);
        assert!(matches!(r, Err(ZipCheckError::TooLarge(_))), "{r:?}");
    }
}
