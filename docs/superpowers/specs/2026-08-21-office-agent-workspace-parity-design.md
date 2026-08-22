# Office Agent Workspace Parity Design

## Goal

Refactor the Office task pane into the same calm, task-oriented Agent workspace used by WisWork PC,
then close the remaining safe capability gaps from the `office-agents` reference. Word, Excel, and
PowerPoint continue to share `@wiswork/agent-core`; only their Office.js skills differ.

The product bet is that an Office Agent is useful when conversation, observable tool work, files,
and approval-gated document changes live in one narrow workspace. A cosmetic chat skin is not
enough: users must be able to follow a multi-turn task, inspect evidence, stop work, review exact
document impact, and recover from failures without leaving Office.

## Non-goals

- No arbitrary JavaScript, `eval`, `Function`, native shell, ambient browser authority, or direct
  model access to raw Office.js objects.
- No simulated success when an Office API set is unavailable.
- No Office access token, Wispaper credential, provider key, or PC filesystem in the task pane.
- No autonomous document mutation. Every write remains a bounded immutable proposal requiring an
  explicit confirmation and semantic post-write verification.
- No change to the WisWork PC Agent harness.

## User experience

The connected task pane becomes a full-height three-region workspace:

1. A compact header shows the active host, Relay/PC connection, session menu, and a new-task action.
2. A scrollable timeline preserves user turns, streamed assistant turns, bounded tool activity,
   errors, attachments, and completion state. Empty state offers host-aware starter prompts.
3. A sticky composer supports multiline text, session attachments, installed skills, send, and
   stop. Attachments and skills open bounded management panels instead of consuming permanent
   vertical space.

Document proposals appear inline at the causal point in the timeline and in a focused review
surface. The review shows title, host, targets, impact count, before/after or structured preview,
verification expectations, Reject, and Confirm. Working and applying are distinct; while applying,
all conflicting actions are disabled. Errors use stable codes plus actionable user-facing copy.

The visual language reuses WisWork PC tokens and interaction hierarchy, adapted for a 280–500 px
task pane. It supports light/dark/high-contrast modes, keyboard operation, reduced motion, and
screen-reader live regions. Raw colors are prohibited outside token definitions.

## Agent and state architecture

`OfficeAgentSession` owns a bounded presentation event log separate from the provider history. It
emits immutable user, assistant, tool, proposal, error, and system events with stable IDs. Streaming
updates replace only the active assistant event. Logout, Relay loss, new task, and runtime disposal
cancel active work and clear conversation, proposals, VFS, installed skills, and pending media.

The existing `AgentLoop` remains the only harness. Host runtime composition remains exactly one of
Word, Excel, or PowerPoint plus shared skills. UI components receive presentation snapshots and
never receive credentials, capabilities, or raw Office.js contexts.

## Capability completion

### Safe conversion worker

`pdf-to-text`, `pdf-to-images`, `docx-to-text`, and `xlsx-to-csv` run in a dedicated terminateable
Web Worker. Inputs come only from the session VFS. Before decompression or rendering, the worker
enforces compressed bytes, entry count, entry path, per-entry and aggregate uncompressed bytes,
page/sheet/row/cell/pixel counts, CPU deadline, output bytes, and cancellation. The main thread
terminates the worker on timeout, logout, or cancellation. Partial outputs are never mounted.

### Web retrieval

`web-search`, `web-fetch`, and `image-search` run only through new fixed PC/Relay capabilities. The
task pane sends a typed bounded request; PC uses the signed-in WisWork account and a fixed server
destination. URLs are HTTPS-only with DNS/IP redirect validation, private-network denial, response
content-type and byte limits, timeout, cancellation, safe error mapping, and no upstream credential
or body leakage. Relay protocol frames remain exact, versioned, bounded, and capability-scoped.

### Import/export and media tools

`csv-to-sheet`, `sheet-to-csv`, `image-to-sheet`, and `insert-image` are declarative host tools, not
shell commands. Reads may execute directly. Writes capture bounded host state, create a structured
proposal, revalidate staleness, execute once, and compare semantic post-state. Tools advertise only
on hosts and API sets where the complete operation and verification are available.

### Skill packages

The UI accepts a bounded package containing exactly one `SKILL.md` and allowlisted auxiliary text
or image assets. Paths are normalized, traversal/symlinks/executables are rejected, total and file
counts are capped, and installation is atomic in the in-memory session registry. Skills contribute
instructions and declarative metadata only; they cannot add executable code or authority.

## Trust and autonomy boundaries

- The Agent may automatically read bounded document/session state and run bounded pure conversion
  or retrieval operations.
- It may recommend a document mutation and prepare a preview.
- Every Office document write and image insertion requires explicit confirmation.
- Network requests are limited to the fixed authenticated retrieval service; arbitrary URLs are
  inputs to that service, not direct task-pane fetch destinations.
- Raw Office.js programs, unsupported API sets, unverifiable writes, unsafe archives, excessive
  files, private-network URLs, and native-shell syntax are refused with stable errors.
- Timeline events expose tool names, bounded summaries, confirmation state, and verification
  outcome so users can see and interrupt Agent work.

## Failure handling and rollback

Each shipped capability has an independent feature flag. The entire workspace can roll back to the
current Office UI without changing Relay identity, document formats, or stored data. Conversion and
skill state is memory-only. Relay/web protocol additions are versioned and ignored by older clients.
Web retrieval remains unshipped until a canonical service is attested; its activation must add the
Office flag and v2 composition atomically after PC/Relay support is deployed.

Timeout, cancellation, logout, offline, malformed output, quota violation, unsupported API, stale
proposal, write failure, and verification failure all fail closed. No error path may claim a write
or retain partial VFS/package output.

## Verification and release gates

- Component tests cover timeline streaming, multi-turn history, composer, panels, keyboard use,
  proposal review, stable errors, dark/high-contrast tokens, and responsive task-pane widths.
- Worker tests use archive bombs, traversal, malformed XML, extreme cell references, oversized
  pages/images, timeout, cancellation, and atomic-output fixtures.
- Relay/PC tests cover exact request schemas, authentication, SSRF/redirect denial, byte/time
  limits, cancellation, logout revocation, and safe errors.
- Every new host write has RED/GREEN tests for preview, stale input, cancellation, exactly-once
  execution, semantic verification, and unsupported API behavior.
- Full Office, agent-core, ai-provider, shell, Relay, typecheck, lint, format, theme, build, license,
  and diff checks pass.
- Manual acceptance runs Word/Excel/PowerPoint on Windows and macOS plus Word Web, including Relay
  pairing, multi-turn chat, each capability family, rejection/confirmation, logout, offline, and
  accessibility smoke checks.
