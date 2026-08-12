# WisWork

**The full-featured open-source AI-native office suite.**

[![License: Apache-2.0](https://img.shields.io/github/license/AtomInnoLab/WisWork)](LICENSE)
[![Latest release](https://img.shields.io/github/v/release/AtomInnoLab/WisWork)](https://github.com/AtomInnoLab/WisWork/releases/latest)
[![Downloads](https://img.shields.io/github/downloads/AtomInnoLab/WisWork/total)](https://github.com/AtomInnoLab/WisWork/releases)
[![GitHub stars](https://img.shields.io/github/stars/AtomInnoLab/WisWork?style=flat)](https://github.com/AtomInnoLab/WisWork/stargazers)
![Platforms: macOS | Windows | Linux](https://img.shields.io/badge/platform-macOS%20%7C%20Windows%20%7C%20Linux-lightgrey)

[Website](https://github.com/AtomInnoLab/WisWork) · [Download](https://github.com/AtomInnoLab/WisWork/releases/latest) · [Demo](https://www.youtube.com/watch?v=B2pLdMX95v4)

WisWork is a free, open-source alternative to Microsoft Office for macOS,
Windows, and Linux, built around AI editing as a first-class workflow rather
than a bolted-on chat box. It opens and saves the real Microsoft Office
formats — Word (`.docx`), Excel (`.xlsx`), PowerPoint (`.pptx`) — and edits
PDF and Markdown too: a word processor, spreadsheet, presentation editor,
PDF editor, and Markdown editor as six Electron apps sharing one engine
layer.

WisWork release artifacts are published through the AtomInnoLab release pipeline.

[![Meet WisWork — the full-featured open-source AI-native office (video)](https://img.youtube.com/vi/B2pLdMX95v4/maxresdefault.jpg)](https://www.youtube.com/watch?v=B2pLdMX95v4)

[Watch the demo video on YouTube](https://www.youtube.com/watch?v=B2pLdMX95v4)

## Features

- **Real PDF editing** — retype text and edit images in the page itself, original fonts preserved.
- **Microsoft Word–compatible, byte-preserving `.docx` editing** — only what you touched changes; Word never notices.
- **Word-faithful pagination** — page breaks land where Word puts them.
- **Excel-compatible spreadsheets** — in-house engine with a Rust `.xlsx` sidecar, own charts, pivot tables, slicers.
- **PowerPoint-compatible presentations** — in-house `.pptx` engine with masters, layouts, smart guides, non-destructive crop.
- **Markdown to Word, fully local** — the same OOXML engine, no Pandoc, no cloud.
- **AI that edits documents** — block-level edits with snapshots and diffs, document-aware agents.
- **Agent tools built in** — web/image search, image generation, media analysis.
- **Light / dark / system themes.**
- **macOS, Windows, Linux.**
- **Free & open-source (Apache-2.0).**

## Download

| Platform                             | Requirements                                          | Download                                                                                                                   |
| ------------------------------------ | ----------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [WisWork-0.6.101-arm64.dmg](https://github.com/AtomInnoLab/WisWork/releases/download/v0.6.101/WisWork-0.6.101-arm64.dmg)   |
| **macOS** — Intel (x64)              | macOS 11+                                             | [WisWork-0.6.101.dmg](https://github.com/AtomInnoLab/WisWork/releases/download/v0.6.101/WisWork-0.6.101.dmg)               |
| **Windows** (x64)                    | Windows 10+                                           | [WisWorkSetup-v0.6.101.exe](https://github.com/AtomInnoLab/WisWork/releases/download/v0.6.101/WisWorkSetup-v0.6.101.exe)   |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 or newer)           | [wiswork_0.6.101_amd64.deb](https://github.com/AtomInnoLab/WisWork/releases/download/v0.6.101/wiswork_0.6.101_amd64.deb)   |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [wiswork-0.6.101.x86_64.rpm](https://github.com/AtomInnoLab/WisWork/releases/download/v0.6.101/wiswork-0.6.101.x86_64.rpm) |
| **Linux** — other distributions      | x86_64, glibc 2.34+, FUSE 2                           | [WisWork-0.6.101.AppImage](https://github.com/AtomInnoLab/WisWork/releases/download/v0.6.101/WisWork-0.6.101.AppImage)     |

All builds come from `main`; the macOS and Windows installers are signed.
Older versions are on the [Releases](https://github.com/AtomInnoLab/WisWork/releases) page.

### Installing on Linux

The deb installs with apt — it pulls in the dependencies and adds WisWork
to the applications menu:

```bash
sudo apt install ./wiswork_0.6.101_amd64.deb
```

On Fedora / RHEL-family / openSUSE, install the rpm instead:

```bash
sudo dnf install ./wiswork-0.6.101.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./wiswork-0.6.101.x86_64.rpm  # openSUSE
```

The AppImage instead runs in place: install the FUSE 2 runtime
(`sudo apt install libfuse2`; on Ubuntu 24.04 the package is `libfuse2t64`),
make the file executable, then run it:

```bash
chmod +x WisWork-0.6.101.AppImage
./WisWork-0.6.101.AppImage
```

## Apps

| App             | Product              | What it is                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| --------------- | -------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `apps/docs`     | **WisWork Docs**     | `.docx` word processor. Byte-preserving round trip: only dirty paragraphs are regenerated (paragraph patch), everything else in the original file is kept byte-for-byte, so opening and saving never breaks layout in Word. Paginated view whose line metrics reproduce the original document's layout, tracked changes, comments, styles, equations, ink.                                                                                                                                                                                                      |
| `apps/sheets`   | **WisWork Sheets**   | `.xlsx` spreadsheet. UI built on the open-source [Univer](https://github.com/dream-num/univer) core (Apache-2.0) with a large layer of in-house extensions; `.xlsx` import/export runs through an in-house Rust sidecar (calamine + IronCalc), charts are rendered in-house (Konva), plus pivot tables, slicers, conditional formatting, and formula tracing.                                                                                                                                                                                                   |
| `apps/slides`   | **WisWork Slides**   | `.pptx` presentations. In-house `.pptx` parse/render/edit engine with masters, charts, cropping, ink, and text shaping (HarfBuzz metrics).                                                                                                                                                                                                                                                                                                                                                                                                                      |
| `apps/pdf`      | **WisWork PDF**      | `.pdf` viewer/editor on [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) + [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT): annotations, forms, outlines, stamps, signatures, page operations, and printing support. True text editing — paragraph selection with in-block reflow, alignment restoration, original-font preservation — and content-stream image insert/edit, all rewriting page content streams through [PDFium](https://pdfium.googlesource.com/pdfium/) wasm (BSD-3-Clause) with subset-embedded fonts — no cover-up annotations. |
| `apps/markdown` | **WisWork Markdown** | `.md` / `.markdown` editor: Tiptap block editor over plain Markdown files — headings, lists, tables, images, code blocks — saved back as plain Markdown, hosted in shell tabs.                                                                                                                                                                                                                                                                                                                                                                                  |
| `apps/latex`    | **WisWork LaTeX**    | AI-native multi-file LaTeX project editor with CodeMirror, confirmed AI proposals, Tectonic compilation, diagnostics, PDF preview, and SyncTeX navigation.                                                                                                                                                                                                                                                                                                                                                                                                      |
| `apps/shell`    | **WisWork**          | The suite shell: home screen, tabbed hosting of all six editors, light/dark/system theme, auto-update.                                                                                                                                                                                                                                                                                                                                                                                                                                                          |

Every app embeds the same AI panel: block-granular AI editing with version
snapshots and diffs in docs, a tool-calling agent over workbook/slide/PDF
state in the others.

**AI service.** The desktop apps authenticate with WisPaper and send model requests through the WisModel-compatible proxy. The service key is read only by the Electron main process from the environment and is never exposed to renderers or stored in project files.

The whole suite ships light / dark / system UI themes built on shared design
tokens (`packages/ui`), with a CI guard that keeps chrome colors on the token
system. Document surfaces stay light in dark mode — Word-style dark chrome
around white paper — so files render and export identically in both themes.

## Engine packages

All pure TypeScript, no Electron dependency, unit-tested (except the UI kit):

- `packages/docx-engine` — docx parsing → block tree (with `docxIndex`
  anchors and passthrough), OOXML fragment generation, byte-level paragraph
  patching.
- `packages/pptx-engine` / `packages/pptx-render` — pptx model and rendering.
- `packages/file-parse` — text extraction for AI attachments (office formats,
  text formats).
- `packages/agent-core` — the AI agent loop and skill composition shared by
  every app.
- `packages/ai-provider` — provider abstraction and streaming for the model
  backends.
- `packages/auth` — WisPaper OAuth, encrypted session storage, refresh, and Electron lifecycle integration.
- `packages/ai-search` — web/image search tools.
- `packages/i18n`, `packages/ui`, `packages/project-store`,
  `packages/electron-utils` — shared i18n core, React UI kit, recent-files
  store, and Electron main-process helpers.

## Development

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all five editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

The sheets app additionally needs a Rust toolchain for its xlsx sidecar
(`cargo` on PATH); `npm run build -w @wiswork/sheets` compiles it
automatically.

LaTeX packaging, Tectonic cache behavior, offline compilation, cleanup, and troubleshooting are
documented in [`docs/development/latex.md`](docs/development/latex.md).

### Development authentication and model access

Run the unified Shell for OAuth development. The Shell is the only process that registers and
owns the `wiswork://oauth/callback` deep link; packaged macOS and Windows builds declare that
protocol in the installer metadata. Docs, Sheets, Slides, and PDF standalone builds deliberately
do not register it and fail closed with `auth_unavailable_in_standalone` when login is requested.

The development OAuth defaults target WisPaper Logto and its Gateway. They can be overridden by
the non-secret main-process environment variables `WISWORK_OAUTH_AUTHORIZATION_URL`,
`WISWORK_OAUTH_CALLBACK_URL`, `WISWORK_OAUTH_REFRESH_URL`, and `WISWORK_OAUTH_CLIENT_ID`.
The desktop redirect URI is `wiswork://oauth/callback`. The login request uses OAuth state and
PKCE S256; the Gateway callback request includes the authorization code, PKCE verifier, desktop
redirect URI, and client ID. Deployment is blocked until the Gateway owner confirms that it
validates and consumes the PKCE verifier for this client. A fixed refresh code, OAuth tokens, or
other service-side credentials must never be configured in the desktop app.

Managed model requests require a valid WisWork login. The main process sends the current OAuth
access token to the fixed WisUsage endpoint `https://wisusage.dev.atominnolab.com/v1/messages`
using the Anthropic Messages protocol and model `qwen/qwen3.8-max`. A 401 refreshes the session
once and retries. Tokens never enter renderer settings or IPC responses, and renderers cannot
override the endpoint, model, authorization header, or required `sg` serving region. Safe failures
expose only the request stage and HTTP status for support diagnostics—never the token or upstream
response body. No model service key is required.

Features formerly supplied by the removed
cloud runtime—image generation, media analysis, cloud slide generation, and cloud PDF conversion—are
unavailable and return `unsupported_feature`; web/image search retains its documented
Serper/DuckDuckGo paths.

Before accepting a real integration build, verify all of the following without recording tokens:

- Gateway contract review confirms PKCE verifier validation and one-time authorization-code use.
- A packaged macOS and Windows build each completes cold-start login, deep-link return, restart
  session restore, refresh, logout, and repeated/expired callback rejection.
- Standalone editor builds neither claim the `wiswork` protocol nor offer a working login path.
- An authenticated Shell streams a model response and completes a tool call against the approved
  development endpoint, while missing credentials fail as `model_credentials_missing`.
- Application logs, renderer IPC traffic, settings, user projects, crash output, and packaged
  resources contain no service key, authorization code, access token, or refresh token.

Local UI/e2e driver scripts (Playwright + Electron, for local acceptance, not
committed by default) live in [`scripts/drivers/`](scripts/drivers/README.md).

## Architecture notes (docx round trip)

```
open docx ─► archive original by hash (never touched)
          ─► docx-engine parses word/document.xml top-level elements (w:p / w:tbl / …)
          ─► Block tree, each block anchored by docxIndex + original XML slice
          ─► Tiptap streaming editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into original document.xml (untouched blocks keep original bytes)
          ─► repack zip; all other entries copied byte-for-byte
```

The same philosophy holds in sheets and slides: the original file is the
source of truth, edits are applied as narrow patches, and everything the
editor didn't touch survives the round trip untouched.

## FAQ

**Is WisWork free?**
Yes. WisWork is free and open-source under the Apache-2.0 license — no
trial, no paid tier for the apps themselves.

**Can WisWork open Microsoft Word, Excel, and PowerPoint files?**
Yes. WisWork opens and saves native `.docx`, `.xlsx`, and `.pptx` files.
Saving is byte-preserving: parts of the file you didn't touch are written
back byte-for-byte, so documents keep working in Microsoft Office.

**Does WisWork work offline?**
Document editing is fully local — files never leave your machine to be
opened, edited, or saved. AI features authenticate through WisPaper and use
WisWork-managed services, so they require a network connection.

**Can WisWork edit PDF files?**
Yes — real PDF text and image editing that rewrites the page content stream
with the original fonts preserved, not cover-up annotations.

## Security

See [SECURITY.md](SECURITY.md) for the process security posture (renderer
sandboxing, IPC validation, external-link gating) and the threat models for
AI-generated content.

## Acknowledgements

WisWork would not be possible without these open-source projects:

- [Electron](https://www.electronjs.org/) — the desktop runtime for every app.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — the spreadsheet
  UI core that Sheets extends.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause, bundled via
  [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)) — the
  content-stream engine behind true PDF text and image editing.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) and
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — PDF rendering and
  document assembly.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  the block editors in Docs and Markdown.
- [Konva](https://konvajs.org/) — canvas rendering for Slides and Sheets
  charts.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — text-shaping
  metrics for complex scripts.
- [calamine](https://github.com/tafia/calamine) and
  [IronCalc](https://github.com/ironcalc/IronCalc) — the read and calc layers
  of the Rust xlsx sidecar.
- Liberation, Carlito, Caladea, and Noto CJK fonts (OFL/Apache-2.0) — bundled
  document fonts.

## Third-party notices

`npm run notices` regenerates the bundled third-party license summary
(`tools/gen-third-party-notices.mjs`); all runtime dependencies are
MIT/Apache-2.0/BSD-3-Clause/OFL, and the bundled fonts (Liberation, Carlito,
Caladea, Noto CJK subsets) are OFL/Apache.

## License

WisWork is licensed under the [Apache License 2.0](LICENSE), with one
exception: the `ee/` directory is reserved for future enterprise modules and
is covered by the [enterprise license](ee/LICENSE).

The WisWork name and logo identify the AtomInnoLab distribution. Apache-2.0 does not grant trademark rights.
