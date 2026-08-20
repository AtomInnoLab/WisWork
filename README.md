# WisWork

An AI-native office suite for macOS and Windows: word processor, spreadsheet,
presentations, and PDF — five Electron apps sharing one engine layer, built
around AI editing as a first-class workflow rather than a bolted-on chat box.

WisWork release artifacts are published through the AtomInnoLab release pipeline.

## Apps

| App                 | Product                       | What it is                                                                                                                                                                                                                                                                                                                                                 |
| ------------------- | ----------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `apps/docs`         | **WisWork Docs**              | `.docx` word processor. Byte-preserving round trip: only dirty paragraphs are regenerated (paragraph patch), everything else in the original file is kept byte-for-byte, so opening and saving never breaks layout in Word. Paginated view whose line metrics reproduce the original document's layout, tracked changes, comments, styles, equations, ink. |
| `apps/sheets`       | **WisWork Sheets**            | `.xlsx` spreadsheet. UI built on the open-source [Univer](https://github.com/dream-num/univer) core (Apache-2.0) with a large layer of in-house extensions; xlsx import/export runs through an in-house Rust sidecar (calamine + IronCalc), charts are rendered in-house (Konva), plus pivot tables, slicers, conditional formatting, and formula tracing. |
| `apps/slides`       | **WisWork Slides**            | `.pptx` presentations. In-house pptx parse/render/edit engine with masters, charts, cropping, ink, and text shaping (HarfBuzz metrics).                                                                                                                                                                                                                    |
| `apps/pdf`          | **WisWork PDF**               | PDF viewer/editor on pdf.js + pdf-lib: annotations, forms, outlines, stamps, signatures, page operations, print.                                                                                                                                                                                                                                           |
| `apps/latex`        | **WisWork LaTeX**             | AI-native multi-file LaTeX project editor with CodeMirror, confirmed AI proposals, Tectonic compilation, diagnostics, PDF preview, and SyncTeX navigation.                                                                                                                                                                                                 |
| `apps/office-addin` | **WisWork Office Add-in Lab** | Office.js development workspace and shared task pane for building Word, Excel, and PowerPoint add-ins.                                                                                                                                                                                                                                                     |
| `apps/shell`        | **WisWork**                   | The suite shell: home screen, tabbed hosting of the editors, auto-update.                                                                                                                                                                                                                                                                                  |

Every app embeds the same AI panel: block-granular AI editing with version
snapshots and diffs in docs, a tool-calling agent over workbook/slide/PDF
state in the others.

**AI service.** The desktop apps authenticate with WisPaper and send model requests through the WisModel-compatible proxy. The service key is read only by the Electron main process from the environment and is never exposed to renderers or stored in project files.

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
npm run dev          # all four editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dev:office   # HTTPS Office.js task-pane development server
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
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
          ─► TipTap streaming editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into original document.xml (untouched blocks keep original bytes)
          ─► repack zip; all other entries copied byte-for-byte
```

The same philosophy holds in sheets and slides: the original file is the
source of truth, edits are applied as narrow patches, and everything the
editor didn't touch survives the round trip untouched.

## Security

See [SECURITY.md](SECURITY.md) for the process security posture (renderer
sandboxing, IPC validation, external-link gating) and the threat models for
AI-generated content.

## Third-party notices

`npm run notices` regenerates the bundled third-party license summary
(`tools/gen-third-party-notices.mjs`); all runtime dependencies are
MIT/Apache-2.0/OFL, and the bundled fonts (Liberation, Carlito, Caladea, Noto
CJK subsets) are OFL/Apache.

## License

WisWork is licensed under the [Apache License 2.0](LICENSE), with one
exception: the `ee/` directory is reserved for future enterprise modules and
is covered by the [enterprise license](ee/LICENSE).

The WisWork name and logo identify the AtomInnoLab distribution. Apache-2.0 does not grant trademark rights.
