# Office Agent Workspace Parity Implementation Plan

## Goal and non-goals

Deliver the approved WisWork-style Office Agent workspace and the remaining safe reference
capabilities. Preserve the existing AgentLoop, Relay identity, confirmation boundary, and
fail-closed host behavior. Do not add arbitrary JavaScript, native shell, direct task-pane network
authority, unverifiable writes, or simulated API support.

## Architecture and global constraints

The task pane renders an immutable presentation timeline produced by `OfficeAgentSession`. Shared
conversion and package capabilities live behind bounded terminateable workers; authenticated web
retrieval is a new typed PC/Relay capability; Office writes stay in host adapters and structured
proposal controllers. All public inputs have exact schemas and hard byte/count/time limits.

Global constraints: no credentials or capabilities in UI/history/logs/storage; no raw colors outside
tokens; no eval/Function/native shell; no write before confirmation; stale/cancel/logout checks
before every irreversible Office operation; semantic verification after every write; exact protocol
frames; safe stable errors; feature-flag rollback.

## Task 1: Conversation workspace and presentation state

Files: `apps/office-addin/src/agent/use-office-agent.ts`, new presentation state module and tests,
`apps/office-addin/src/App.tsx`, Office styles/tokens, UI tests, README.

Deliverable: bounded multi-turn user/assistant/tool/proposal/error timeline; streaming replacement;
new-task/logout clearing; full-height responsive header/timeline/sticky composer; host-aware empty
state; attachment/skill panels; inline proposal review; accessible working/applying/stop/retry
states using WisWork tokens.

TDD: first add failing state tests for two turns, streamed updates, tool lifecycle, new-task/logout,
and proposal placement, plus component tests for keyboard/composer/panels and token checks. Implement
the state projection, then components/styles. Run Office targeted/full tests, typecheck, build,
lint, format, theme, and screenshot/manual responsive inspection. Scoped commit.

## Task 2: Terminateable conversion worker

Files: new `apps/office-addin/src/skills/shared/conversion-worker*`, conversion schemas/fixtures/tests,
`commands.ts`, shared skill registry, Vite worker build configuration, README/dependency audit.

Deliverable: real `pdf-to-text`, `pdf-to-images`, `docx-to-text`, and `xlsx-to-csv` commands with
atomic VFS outputs and hard archive/page/sheet/cell/pixel/output/deadline limits. Main-thread timeout
and cancellation terminate the worker.

TDD: add failing happy-path and adversarial fixtures before implementation: ZIP bomb metadata,
traversal, malformed XML, extreme A1 coordinate, oversized image/page, hanging worker, cancellation,
and no partial output. Run shared/Office suites, typecheck/build, dependency/license/CSP/bundle audit,
lint/format/diff. Scoped commit.

## Task 3: Authenticated web capability through PC and Relay

Files: Relay protocol/service tests, PC Relay client and fixed retrieval proxy, shared IPC/preload only
if needed for the trusted renderer, Office Relay session/transport, shared web skill, deployment docs.

Deliverable: versioned `web-search`, `web-fetch`, and `image-search` requests using the signed-in PC
identity and fixed service destination, with streaming/bounded responses, cancellation and logout
revocation. Office advertises the tools only after protocol capability negotiation.

TDD: failing cross-component protocol tests for exact schemas, unauthenticated use, wrong capability,
private/link-local/loopback URL, redirect-to-private, unsupported content type, oversized response,
timeout, cancel, duplicate IDs, disconnect, and safe errors. Implement Relay routing and PC proxy,
then Office client/tools. Run Rust, shell, Office and security checks. Deploy PC/Relay before enabling
Office flag. Scoped commits per independently deployable protocol/server/client unit.

## Task 4: Declarative import/export and media tools

Files: shared schemas; Excel/Word/PowerPoint skill and browser adapters; host tests; README inventory.

Deliverable: `csv-to-sheet`, `sheet-to-csv`, `image-to-sheet`, and `insert-image` on only the hosts
and API sets that support bounded execution and semantic verification. Image bytes originate from
bounded VFS media; CSV parsing is streaming/bounded and formula-injection-safe on export.

TDD: failing tests for exact schemas, dimensional limits, hostile CSV, image limits/types, stale
state, reject, cancellation, partial host failure, idempotence, and post-write mismatch. Implement
reads, then proposal-gated writes. Run host/full Office tests and Windows/macOS manual API-set matrix.
Scoped commit.

## Task 5: Bounded multi-file skill packages

Files: skill package parser/registry/VFS, upload UI, runtime composition, tests and docs.

Deliverable: atomic installation/removal of a package with one `SKILL.md` plus bounded allowlisted
text/images. Package content can alter Agent instructions and provide VFS resources but cannot add
code, commands, network destinations, or Office authority.

TDD: failing tests for traversal, duplicate/case-colliding paths, symlink/executable entries, missing
or duplicate manifest, count/size/decompression limits, invalid UTF-8, cancellation, duplicate
package, atomic rollback, uninstall and logout cleanup. Implement parser and UI management panel.
Run Office/full security checks. Scoped commit.

## Task 6: Integration, release and rollback

Files: manifests/build config, capability flags, deployment/runbooks, manual acceptance checklist,
version notes.

Deliverable: capability negotiation prevents older PC/Relay clients from seeing unavailable tools;
new workspace and each shipped capability family have independent rollback flags; the blocked Web
family receives its flag only in the same change that activates an attested service and Office v2
composition. Deployment order and diagnostics are documented.

Verification: fresh full repository tests; Office/agent-core/ai-provider/shell/Relay typechecks;
Rust tests/clippy/deny; lint, format, theme, diff and configured/unconfigured production builds;
artifact scan for credentials, eval, native shell, wildcard origins, source maps and raw colors.
Perform broad independent review and fix Critical/Important findings. Manual Windows/macOS/Web
acceptance remains an explicit release gate. Final scoped release commit and branch handoff.
