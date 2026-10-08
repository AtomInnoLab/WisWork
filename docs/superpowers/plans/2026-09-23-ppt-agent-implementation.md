# PPT Agent implementation ledger

Approved specification: `docs/product/wiswork-ppt-agent-solution-and-implementation-plan-2026-09-22.md` (v1.3).
Baseline: `482dedf`, isolated branch `codex/ppt-agent-implementation`. The baseline preserves existing uncommitted ACP changes; those are not newly implemented by this task.

## Architecture and constraints

Agent-generated, validated SlideIR is compiled by a deterministic PptxGenJS compiler. A PC service persists input, compiled artifacts and idempotent receipts; a negotiated `presentation.v1` capability exposes this to PowerPoint. Existing Office.js and controlled OOXML tools continue handling local edits. The generation skill returns a downloadable editable PPTX plus honest QA results; subsequent import must retain document binding and write protection.

Preserve original working tree. Never evaluate model JavaScript, fetch arbitrary asset paths, fabricate evidence, claim visual checks without screenshots, or overwrite user content during generation. Retain existing per-document relay pool rather than rewriting it. Document session identity must be derived from the host, not supplied by the model. No deployment is part of a local implementation result.

## Baseline findings (stage 0)

- PC production currently uses `plan_deck` and native per-object tools; `build_deck` is absent.
- PptxGenJS is only a test dependency. Shared SlideIR and durable PresentationProject do not yet exist.
- Three-client connection isolation already has tests. Stale socket messages can still revoke newer pairings.
- Screenshots already reach the visual model. The legacy review gate described by the specification is absent.
- Existing Office proposals already validate baselines, verify writes and offer controlled recovery.

## Independently reviewable deliverables

1. **Connection isolation and compilation transport**
   - Files: addin `relay/session.ts` and tests; shell `office-relay-client.ts` and tests; relay Rust capability allowlist/tests.
   - Accept: stale generations cannot affect active pairings; compilation advertised only with installed handler; negotiated requests preserve auth, cancellation, bounds and old-client compatibility.
   - RED: late old-socket message or old proxy completion affects new pairing; presentation request unsupported. GREEN: dedicated regression and complete relevant suites.
   - Commit: `feat(office): route negotiated presentation compilation safely`.
2. **Shared IR and native compiler**
   - Files: `packages/pptx-engine/src/presentation.ts`, compiler, package exports and tests/fixtures.
   - Interfaces: browser-safe parse/schema/geometry API; `compilePresentationDeck(unknown)` returns bytes and explicit check statuses.
   - Accept: native text/shape/image/table/chart output; invalid input rejected; real 8-page PPTX reopens with expected content; no false visual/source validation.
   - Commit: `feat(presentation): compile validated slide IR with PptxGenJS`.
3. **Durable projects and receipts**
   - Files: project-store persistence module/tests; shell presentation service/tests.
   - Interfaces: compile/get requests bound to document ID and opaque project/request identifiers.
   - Accept: identical retries after recreation do not recompile; conflicting keys rejected; prior ready output survives failure; atomic bounded records; cancellation and concurrent retries verified.
   - Commit: `feat(presentation): persist compilation projects and idempotent receipts`.
4. **Taskpane generation and delivery**
   - Files: addin generation skill/tests, host runtime/tests, App download/recovery entry, document identity helper/tests; shell index wiring.
   - Accept: only negotiated PowerPoint session exposes compiler; validated input goes to PC; output reaches existing VFS and user download; source/report accompanies PPTX; restore binds host identity and project ID. Test composed runtime and failure paths, build addin/shell.
   - Commit: `feat(office): expose recoverable presentation generation in Taskpane`.

## Verification and stage exit

Run failing behavior tests before implementation; then full relevant package suites, full repository `npm test` and `npm run typecheck`, addin/shell production builds, and Rust relay tests. Run independent review of implemented units and final integration. Report environment failures and existing baseline failures separately. Use generated PPTX round-trip tests as structural evidence only; real Office host rendering/editability remains a separate gate.

The complete approved roadmap remains the objective. This ledger's first delivery closes the missing production path. O0 automatic secure reconnect, O1 full project timeline, O2 resumable large asset transport, O3 per-page host import/transactions, O4 visual/round-trip gates, O5 change-set UX and O6 compatibility release must be reported individually, not inferred from successful file compilation. P1/P2 follow their specification dependencies. Real 20-task professional benchmark and 16/20 acceptance are not satisfied by synthetic fixtures.

## Rollback, migration and release

Additive, versioned project records only; reject unsupported versions without mutation. Capability negotiation hides generation on older servers/PCs. Removing the generation handler disables advertisement without changing existing tools. Preserve user files and committed receipts on failure. Relay, PC and addin must all carry compatible capability support before release; do not claim deployed service support from local tests. No automatic merge or release.
