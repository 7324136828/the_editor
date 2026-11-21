# Office Editor: Chrome / Edge extension

This branch holds the Office Editor browser extension (Manifest V3). It lets you
view and edit Word, Excel, PowerPoint and LaTeX documents (`.docx`, `.doc`,
`.rtf`, `.odt`, `.md`, `.txt`, `.xlsx`, `.xlsm`, `.xls`, `.ods`, `.csv`, `.tsv`,
`.pptx`, `.tex`) in Chrome and Edge, without Office. The branch also includes
the optional local Rust backend, `office-backend`.

| Path | What it is |
|---|---|
| [`office-chrome-extension/`](office-chrome-extension/) | The extension: editors, popup, settings, link/download integration, tests. **Full documentation:** [`office-chrome-extension/README.md`](office-chrome-extension/README.md) |
| [`src/`](src/) | `office-backend`, the optional local Rust service on `127.0.0.1:47831`. It adds `.doc`/`.rtf`/`.odt`, higher-fidelity Office output and LaTeX → PDF. Start it with `src/run-backend.bat`, `src/run-backend.ps1` or `src/run-backend.sh`. |
| [`example/`](example/) | Sample documents used by the tests and for trying the editor. |
| [`.github/workflows/chrome-extension.yml`](.github/workflows/chrome-extension.yml) | CI: tests, the packed extension (`.crx`) and `.zip`, and releases. |

## Quick start

1. Open `chrome://extensions` (or `edge://extensions`) and turn on **Developer mode**.
2. Click **Load unpacked** and select `office-chrome-extension/`. You can also
   download the `office-editor-extension-*` artifact from a CI run, unzip
   `office-editor-<version>.zip`, and load that folder.
3. Optional: start the backend with `src/run-backend.bat` (Windows) or
   `./src/run-backend.sh` (macOS/Linux). Building it needs Rust from <https://rustup.rs>.

## Build the packed extension locally

```sh
cd office-chrome-extension
node tools/build.mjs --crx     # dist/extension.zip + dist/extension.crx (+ dist/build-info.json)
```

The first `--crx` build creates `office-chrome-extension/key.pem`. That key
fixes the extension ID, so keep it safe and never commit it (`*.pem` is
git-ignored).

## Continuous integration

[`.github/workflows/chrome-extension.yml`](.github/workflows/chrome-extension.yml) runs on every push to
`main` / `chrome-extension`, on pull requests, and on manual runs:

| Job | What it does |
|---|---|
| Extension unit tests | `npm test` in `office-chrome-extension/` (Node 24) |
| Browser and end-to-end tests | The real extension UI in headless Edge on Windows: Word DOM tests, then the Word, Excel, PowerPoint and LaTeX e2e suites. Saved files are validated with openpyxl and python-pptx. This job does not block packaging, because some checks have timing budgets. |
| Backend tests | `cargo test --locked` in `src/` |
| Build packed extension | Builds and verifies a signed `office-editor-<version>.crx` plus `office-editor-<version>.zip`, and uploads them as the `office-editor-extension-<version>` artifact with `SHA256SUMS.txt` and `build-info.json` (version and extension ID) |
| Backend binary | Release builds of `office-backend` for Linux x64, Windows x64 and macOS arm64 (tags and manual runs only) |
| Publish GitHub release | For tags `v*` that match the manifest version (for example `v0.3.0`): creates a GitHub release with the `.crx`, the `.zip` and the backend binaries attached |

### Signing key (recommended)

The `.crx` is signed with the `CRX_PRIVATE_KEY` repository secret. Without the
secret, CI signs with a throwaway key, and the extension ID changes on every
build. To set the secret:

```sh
cd office-chrome-extension
node tools/build.mjs --crx                      # creates key.pem if it does not exist
gh secret set CRX_PRIVATE_KEY < key.pem         # or: Settings → Secrets and variables → Actions
```

### Releasing

1. Bump `version` in both `office-chrome-extension/manifest.json` and
   `office-chrome-extension/package.json`. The build fails if they differ.
2. Tag and push: `git tag v0.3.1 && git push origin v0.3.1`.

Chrome on Windows and macOS installs `.crx` files only from the Chrome Web
Store, or through enterprise policy (`ExtensionInstallForcelist` with an
update URL). For the store, upload the `.zip`. For managed deployment, use the
`.crx`. For development, use **Load unpacked**.

## Tests

```sh
cd office-chrome-extension && npm test     # unit tests (no browser needed)
node tests/browser/run.mjs                 # Word editor DOM tests in headless Edge/Chrome
node tests/e2e/run.mjs                     # end-to-end UI tests (see the extension README for all suites)
cd ../src && cargo test                    # backend tests
```

## License

See [LICENSE](LICENSE).
