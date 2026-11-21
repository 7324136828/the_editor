#!/usr/bin/env sh
# Start the Office Editor backend (office-backend) on http://127.0.0.1:47831.
# Usage: ./run-backend.sh [--port PORT]
# Keep the terminal open while you use the extension; Ctrl+C stops it.
set -e
cd "$(dirname "$0")"

if command -v cargo >/dev/null 2>&1; then
  echo "Building and starting office-backend (the first build takes a few minutes)..."
  exec cargo run --release -p office-backend -- "$@"
elif [ -x target/release/office-backend ]; then
  echo "cargo not found - starting the previously built target/release/office-backend"
  exec target/release/office-backend "$@"
else
  echo
  echo "Rust is not installed, so office-backend cannot be built."
  echo " 1. Install Rust from https://rustup.rs"
  echo " 2. Open a new terminal and run this script again."
  echo
  echo "The extension still works without the backend for .docx, .xlsx, .xls, .ods,"
  echo ".csv, .tsv, .md and .txt files; .doc, .rtf and .odt need the backend."
  exit 1
fi
