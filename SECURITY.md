# Security Policy

## Reporting a Vulnerability

Please report suspected vulnerabilities privately via GitHub's
[private vulnerability reporting](https://github.com/atominnolab/wiswork/security/advisories/new)
on this repository. Do not open public issues for security reports. We aim to
acknowledge reports within 72 hours.

## Process Security Posture

All application windows run with the full Electron renderer lockdown:

- `contextIsolation: true`, `nodeIntegration: false`, `sandbox: true` for every
  document window and tab view (docs, sheets, slides, pdf, shell, updater).
- Renderers reach the main process only through typed, validated IPC channels
  (payloads are schema-checked in the main process; sheets uses zod end to end).
- Every `shell.openExternal` call goes through a single shared gate
  (`@wiswork/electron-utils` → `safeExternalUrl`) that parses the URL and
  enforces a protocol allowlist (http/https; pdf link annotations additionally
  allow mailto). `file:`, `javascript:`, and custom schemes are always rejected.
- No API keys are hardcoded. `WISWORK_MODEL_API_KEY` is read only in the Electron main process and is never returned over IPC or persisted by the app.

## OAuth and Model Credential Boundaries

- Only the unified WisWork Shell owns `wiswork://oauth/callback`. Standalone editor builds do not
  register the protocol and login fails closed with `auth_unavailable_in_standalone`.
- OAuth authorization uses high-entropy state and PKCE S256. Callback URLs are parsed against an
  exact scheme/host/path and bounded field allowlist; pending transactions expire, have bounded
  capacity, and are consumed once. Sessions are persisted only through Electron `safeStorage`.
  Linux backends that provide plaintext or unknown storage are rejected.
- The Gateway callback contract must validate and consume the PKCE verifier supplied by the
  desktop client. A deployment must not ship based only on client-side PKCE generation; the real
  Gateway behavior is a release acceptance gate.
- OAuth access and refresh tokens never cross preload IPC. Authentication error events expose only
  stable categories, and callback values, token bodies, and credentials must not be logged.
- `WISWORK_MODEL_API_KEY` is a development service credential, not a user token. It is accepted
  only from the Electron main-process environment, is not persisted, and cannot be overridden by
  a renderer-provided URL, header, or settings value. Production packaging must not embed it.
- Model errors may record the provider, model, HTTP status, and bounded non-sensitive diagnostics;
  they must never include authorization headers, request credentials, OAuth tokens, authorization
  codes, or raw authentication responses.
- The current model path is a direct development proxy. WisUsage metering, user billing, Gateway
  model forwarding, production key distribution, and server-side rate-limit policy are not
  implemented by this repository and must not be inferred from a successful desktop login.

Image generation, media analysis, cloud slide generation, and cloud PDF conversion are explicitly
unsupported in the current build and return `unsupported_feature`. They do not silently fall back
to a renderer-supplied API key or an unapproved external provider.

## Threat Model: AI-Generated Layout Scripts (slides)

The slides AI can adjust slide layouts by emitting a small script that is
parsed with Acorn and evaluated by a constrained AST interpreter
(`apps/slides/src/renderer/ai/layout-script-interpreter.ts`). The source looks
like a small, synchronous subset of JavaScript for model compatibility, but it
is not passed to `eval`, `Function`, a VM context, a worker, or the JavaScript
engine as executable source.

**What the script can do by design:** read prototype-free JSON copies of
`els`/`canvas`, perform bounded arithmetic/control flow, use explicitly
implemented string/array/regular-expression/Math helpers, and call
`setBox/moveBy/resizeBy/setText/setStyle/setFill/setStroke/log`. Every edit
primitive validates its arguments (element existence, read-only flags, finite
numbers, hex colors) and writes only into an op buffer that is applied through
the same command pipeline as manual edits.

**Interpreter boundary:**

1. Identifiers resolve only in interpreter-owned lexical scopes seeded with the
   documented data and callables. There are no ambient globals, module loader,
   DOM, network, IPC bridge, timers, process APIs, or dynamic code primitives.
2. Property reads are dispatched by value type. Data objects expose own JSON
   fields only; arrays, strings, and regexes expose a small method allowlist.
   Host prototypes and function properties are never traversed, including
   through computed property names.
3. Calls accept only interpreter-created functions or explicit builtins. A host
   function obtained through a constructor/prototype chain cannot be
   represented.
4. Inputs and values crossing into edit primitives are recursively copied as
   JSON-like, prototype-free data. Errors discard all buffered operations;
   logs are capped.
5. Execution has statement/expression and call-depth limits to bound runaway
   loops or recursion.

The Electron renderer sandbox remains defense in depth, but it is not the
layout-script security boundary. The interpreter is designed so a layout
script cannot obtain renderer capabilities in the first place.

If you find a way for a layout script to reach anything beyond the injected
primitives (network, storage, IPC channels not reachable by design, or the
main process), that is a vulnerability — please report it.

## Threat Model: Rendering AI-Generated HTML (slides export)

The HTML-to-pptx export pipeline renders AI-generated HTML in a hidden
`BrowserWindow`. That window is treated as hostile content: full renderer
lockdown (`sandbox: true`, `contextIsolation: true`, `nodeIntegration: false`),
no preload script, no IPC surface — the main process drives it exclusively
through `executeJavaScript` and destroys it under a watchdog timeout.

## Out of Scope

- The cloud AI services this client talks to are operated separately and are
  not part of this repository; issues with them should be reported through the
  service provider's channels.
- Vulnerabilities that require an already-compromised machine or a modified binary. This includes deliberate environment-variable overrides for local development; controlling the process environment is equivalent to code execution on the machine.
