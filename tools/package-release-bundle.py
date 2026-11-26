#!/usr/bin/env python3
"""Build one distributable zip containing the unpacked extension and backend binaries.

The expected backend binary names are the names produced by the GitHub Actions
backend-binaries job, for example office-backend-windows-x64.exe.
"""

import argparse
import hashlib
import json
import pathlib
import shutil
import sys
import zipfile

ZIP_DATE = (2020, 1, 1, 0, 0, 0)
BACKEND_BINARIES = {
    "linux-x64": ("office-backend-linux-x64", "office-backend"),
    "windows-x64": ("office-backend-windows-x64.exe", "office-backend.exe"),
    "macos-arm64": ("office-backend-macos-arm64", "office-backend"),
}


def zip_write_file(bundle, source, arcname, executable=False):
    info = zipfile.ZipInfo(arcname, ZIP_DATE)
    info.compress_type = zipfile.ZIP_DEFLATED
    mode = 0o755 if executable else 0o644
    info.external_attr = (mode & 0xFFFF) << 16
    with open(source, "rb") as src, bundle.open(info, "w") as dst:
        shutil.copyfileobj(src, dst, length=1024 * 1024)


def safe_member_name(name):
    path = pathlib.PurePosixPath(name)
    if path.is_absolute() or ".." in path.parts:
        raise ValueError(f"unsafe path in extension zip: {name}")
    return path.as_posix()


def extension_files(extension_zip):
    with zipfile.ZipFile(extension_zip) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = safe_member_name(info.filename)
            if pathlib.PurePosixPath(name).name in {".DS_Store", "Thumbs.db", "desktop.ini"}:
                continue
            yield name, zf.read(info)


def readme_text(version, platforms):
    backend_lines = []
    chmod_lines = []
    for platform in platforms:
        _, binary_name = BACKEND_BINARIES[platform]
        backend_lines.append(f"backend/{platform}/{binary_name}")
        if not binary_name.endswith(".exe"):
            chmod_lines.append(f"  chmod +x backend/{platform}/{binary_name}")

    contents = "\n".join(backend_lines)
    chmod = "\n".join(chmod_lines)
    chmod_block = ""
    if chmod:
        chmod_block = f"""
On Linux or macOS, make the backend executable if needed:
{chmod}
"""

    return f"""Office Editor {version} - extension and backend bundle
=====================================================

This bundle contains the Office Editor browser extension and the optional
office-backend binaries in one zip file.

Contents
--------
extension/    unpacked Chrome/Edge extension
{contents}

Install
-------
1. Extract this zip.
2. Open chrome://extensions or edge://extensions, enable Developer mode, choose
   "Load unpacked", and select this bundle's extension/ directory.
3. Optional but recommended: start the backend binary for your operating system
   and keep it running while using the editor. The backend adds .doc/.rtf/.odt
   support, higher-fidelity Office output, and LaTeX to PDF conversion.
{chmod_block}
The unsigned backend binaries are intended for local use. Your operating system
may ask for confirmation before running them.
"""


def build_bundle(version, extension_zip, backend_dir, output, platforms=None):
    extension_zip = pathlib.Path(extension_zip)
    backend_dir = pathlib.Path(backend_dir)
    output = pathlib.Path(output)

    if not extension_zip.is_file():
        raise FileNotFoundError(f"extension zip not found: {extension_zip}")

    platforms = list(platforms or BACKEND_BINARIES)
    backends = []
    missing = []
    for platform in platforms:
        artifact_name, binary_name = BACKEND_BINARIES[platform]
        source = backend_dir / artifact_name
        if not source.is_file():
            missing.append(artifact_name)
            continue
        backends.append((platform, source, binary_name))
    if missing:
        raise FileNotFoundError("missing backend binaries: " + ", ".join(missing))

    output.parent.mkdir(parents=True, exist_ok=True)
    manifest_entries = []
    with zipfile.ZipFile(output, "w") as bundle:
        readme = readme_text(version, platforms).encode("utf-8")
        info = zipfile.ZipInfo("README.txt", ZIP_DATE)
        info.compress_type = zipfile.ZIP_DEFLATED
        info.external_attr = 0o644 << 16
        bundle.writestr(info, readme)
        manifest_entries.append({"path": "README.txt", "sha256": hashlib.sha256(readme).hexdigest()})

        extension_count = 0
        for name, data in extension_files(extension_zip):
            arcname = f"extension/{name}"
            zinfo = zipfile.ZipInfo(arcname, ZIP_DATE)
            zinfo.compress_type = zipfile.ZIP_DEFLATED
            zinfo.external_attr = 0o644 << 16
            bundle.writestr(zinfo, data)
            extension_count += 1

        for platform, source, binary_name in backends:
            arcname = f"backend/{platform}/{binary_name}"
            zip_write_file(bundle, source, arcname, executable=True)
            digest = hashlib.sha256(source.read_bytes()).hexdigest()
            manifest_entries.append({"path": arcname, "sha256": digest, "platform": platform})

        bundle_info = {
            "name": "office-editor",
            "version": version,
            "extensionFiles": extension_count,
            "backendPlatforms": [platform for platform, _, _ in backends],
            "files": manifest_entries,
        }
        encoded = (json.dumps(bundle_info, indent=2) + "\n").encode("utf-8")
        info = zipfile.ZipInfo("bundle-info.json", ZIP_DATE)
        info.compress_type = zipfile.ZIP_DEFLATED
        info.external_attr = 0o644 << 16
        bundle.writestr(info, encoded)

    return extension_count, backends


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--version", required=True, help="extension/version used in bundle documentation")
    parser.add_argument("--extension-zip", required=True, help="built extension zip to unpack into the bundle")
    parser.add_argument("--backend-dir", required=True, help="directory containing the CI backend binary artifacts")
    parser.add_argument("--output", required=True, help="path of the combined zip to write")
    parser.add_argument(
        "--platform",
        action="append",
        choices=list(BACKEND_BINARIES),
        dest="platforms",
        help="backend platform to include; repeat for multiple platforms (default: all)",
    )
    args = parser.parse_args(argv)

    try:
        platforms = list(dict.fromkeys(args.platforms or BACKEND_BINARIES))
        count, backends = build_bundle(args.version, args.extension_zip, args.backend_dir, args.output, platforms)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        print(f"bundle failed: {exc}", file=sys.stderr)
        return 1

    platforms = ", ".join(platform for platform, _, _ in backends)
    size = pathlib.Path(args.output).stat().st_size
    print(f"Wrote {args.output} ({size} bytes): {count} extension files, backends: {platforms}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
