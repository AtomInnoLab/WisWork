# Office Safe Diagnostics and Markdown Implementation Plan

## Goal and constraints

Implement the approved safe diagnostic pipeline and taskpane Markdown rendering. Preserve the privacy exclusions, exact protocol schema, bounded memory/wire sizes, best-effort semantics, capability authentication, and independent rollback flags defined in the design.

## Files and responsibilities

- `services/wiswork-relay/src/lib.rs`, `services/wiswork-relay/tests/relay.rs`: validate/authenticate/rate-limit/log `office.diagnostic`.
- `apps/office-addin/src/relay/session.ts`: encode bounded diagnostics over an approved Relay v2 session without affecting Agent requests.
- `apps/office-addin/src/diagnostics/*`: safe event model, sanitizer, ring buffer, lifecycle correlation, and export.
- `apps/office-addin/src/agent/*`, host adapters: record safe tool/proposal/write/verify failures at the closest phase boundary.
- `apps/office-addin/src/App.tsx`, styles and tests: Copy diagnostics action and safe assistant Markdown.
- `packages/ui/src/Markdown.tsx` and tests if needed: streaming-safe Markdown subset and explicit unsupported-content behavior.
- Office build config/docs: exact remote-diagnostics rollback flag and operational documentation.

## Task 1: Relay diagnostic protocol

Create RED Rust integration tests for a valid capability-bound event, invalid session/capability, unknown/extra fields, prohibited/unbounded values, event/session caps, and proof that PC receives nothing. Implement the exact v2 frame validation, per-session counter, best-effort acknowledgement or stable error, and structured safe logging. Run focused tests, Relay full tests, clippy and cargo-deny; commit only Relay scope.

Acceptance: a valid sanitized failure is logged with event/trace/session correlation; no secret or arbitrary user field is accepted; diagnostics cannot disturb the active request or PC connection.

## Task 2: Taskpane diagnostics and export

Create RED tests for the 200-event ring, sanitizer exclusions, error identifier extraction, run/tool/phase correlation, logout/dispose clearing, 4 KiB wire bound, disabled upload, upload failure isolation, and clipboard export. Add the diagnostic collector, instrument common Agent/proposal phase boundaries, extend the Relay session with a best-effort diagnostic sender, and expose a narrow UI export callback. Document the flag and event lookup workflow. Run focused and full Office checks; commit Office diagnostic scope.

Acceptance: a Word/Excel write or verification failure produces a locally exportable trace and, when enabled and paired, one matching sanitized Relay event without document content.

## Task 3: Safe Markdown rendering

Create RED component/renderer tests proving assistant Markdown renders headings/lists/emphasis/code while raw HTML, links/images and incomplete streaming syntax cannot create HTML/network authority. Integrate the shared renderer only for assistant timeline events and add taskpane-specific token-based styles. Run focused UI tests, Office full tests, typecheck, theme and production build; commit Markdown scope.

Acceptance: streamed and completed assistant messages render the allowed subset; user/error/tool/proposal content stays plain text; no unsafe URL or HTML node is created.

## Task 4: Integration, security review, and release

Run fresh full Relay and Office verification plus root lint/format/theme/license/diff gates. Inspect the built taskpane CSP and bundle for no new network destinations, raw HTML execution, source maps, or prohibited diagnostic keys. Independently review the complete diff for protocol, privacy, lifecycle, and Markdown injection issues; fix Critical/Important findings with targeted RED/GREEN tests. Deploy Relay first, then taskpane with remote diagnostics enabled; verify a synthetic safe event and perform manual Word/Excel failure correlation. Roll back by disabling the taskpane flag before reverting Relay support.

Acceptance: automated gates pass, review has no remaining Critical/Important findings, and a manual trace can be correlated without exposing document data.
