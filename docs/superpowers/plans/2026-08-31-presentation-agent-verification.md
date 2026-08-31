# Presentation Agent Planning and Verification Implementation Plan

## Goal and non-goals

Implement the approved task-planning and task-specific verification loop in shared Agent Core, WisWork PC Slides, and PowerPoint Taskpane. Do not broaden editing capability families, alter document formats, or weaken existing transaction/proposal confirmation and recovery boundaries.

## Architecture

A shared strict contract package defines acceptance plans, visual-review results, and completion receipts. Agent Core coordinates clarify/plan/confirm and requires a terminal receipt for presentation mutations. Each host supplies an authority adapter for inspection, mutation receipts, rendering, deterministic verification, and bounded corrections.

## Global constraints

- At most 50 checks, 8 screenshots per pass, 2 MiB visual request, and two corrections.
- No raw document content, paths, IDs, screenshots, prompts, or credentials in diagnostics/receipts.
- No mutation outside the frozen contract.
- Applied-but-unverified and uncertain results are never retried as fresh mutations.
- Existing host confirmation, lease, stale, recovery, undo, and cancellation semantics remain authoritative.

## Task 1 — Shared presentation verification contracts

**Files:** create `packages/presentation-verification/`; update root manifests/lock/scripts.

Implement strict array/object parsers, bounds, canonical contract digest, safe result/receipt schemas, and deterministic receipt rendering inputs. Add parser and adversarial descriptor/prototype/accessor tests.

**Acceptance:** malformed/overbound/unknown inputs fail closed; canonical equivalent contracts share a digest; receipts contain no raw content; workspace test/typecheck/lint pass.

**Commit:** `feat(agent-core): define presentation verification contracts`.

## Task 2 — Agent Core planning and completion orchestration

**Files:** `packages/agent-core/src/{skill,loop}.ts`, exports, tests.

Add optional presentation planning/review hooks, bounded clarify/plan events, a terminal completion receipt boundary, cancellation-safe corrective turns, and receipt-derived final response facts. Preserve non-presentation skill behavior.

TDD cases: simple task bypasses question; ambiguous task asks once; compound task emits plan; rejected confirmation causes zero dispatch; applied-unverified is not retried; cancellation before/after dispatch; no terminal success without receipt.

**Acceptance:** existing Agent Core tests plus new state-machine tests pass; provider history remains paired; hooks fail open only before mutation and fail closed after dispatch where truth is unknown.

**Commit:** `feat(agent-core): orchestrate verified presentation tasks`.

## Task 3 — PC Slides acceptance compiler and deterministic verifier

**Files:** create renderer/shared verification modules; extend `slides-skill.ts`, transaction access types, tests.

Compile supported user intents/plans into contracts for text, style/color, geometry, fill/stroke, and background. Resolve roles/targets through authoritative durable identities and reference-page tokens. Verify post-write properties and tolerances against the current revision.

TDD fixtures include the reported pages 6–8 title/body/emphasis colors and page-6 title reference geometry, ambiguous reference, locked target, no-op, stale revision, and unauthorized expansion.

**Acceptance:** exact supported requests produce deterministic checks; ambiguity requests clarification; stale/unsupported checks never pass; no mutation is introduced by inspection.

**Commit:** `feat(slides): compile task acceptance contracts`.

## Task 4 — PC rendered verification and bounded corrections

**Files:** `AiPanel.tsx`, `slide-qc.ts`, new task-review module, IPC/access tests.

Capture authoritative affected/reference pages after refresh, run strict task-specific visual review, validate fix intents, execute at most two low-risk canonical correction transactions, and publish the completion receipt. Keep generic QC separate and subsequent.

TDD cases: screenshot revision mismatch, visual pass, one correction pass, two-pass stop, screenshot unavailable after applied write, unsafe fix intent, cancellation, session switch, rollback/history truth, and final receipt-derived chat text.

**Acceptance:** no free-form success before receipt; correction touches only failed checks; applied-unverified never reexecutes; PC full tests/typecheck/build pass.

**Commit:** `feat(slides): verify edits against rendered results`.

## Task 5 — PowerPoint Taskpane authority and deterministic verification

**Files:** PowerPoint skill/adapter, proposal/session interfaces, tests.

Add the shared contract adapter, bind contracts to Office session/document/proposal hashes, expose safe post-write property checks, and return normalized mutation receipts. Reuse existing bounded convergence and confirmation flow.

TDD cases: text/style/geometry/background checks, stale proposal, Office lag, post-dispatch cancellation, master/chart/table/XML risk classification, and proved applied versus uncertain.

**Acceptance:** existing PowerPoint transaction safety is unchanged; unsupported checks are unavailable; elevated-risk edits never become automatic corrections.

**Commit:** `feat(office): verify PowerPoint edit contracts`.

## Task 6 — PowerPoint post-write screenshot review and corrections

**Files:** PowerPoint skill/session orchestration, screenshot adapter, Agent session/UI tests.

Capture affected/reference slides after proved writes, bind screenshots to current authority, call the shared visual-review schema, execute up to two allowed native corrections through proposals, and render completion receipts in Taskpane.

TDD cases mirror PC golden tasks plus screenshot failure, readback lag, proposal confirmation, PowerPoint Mac capability gates, session replacement, cancellation, and no Relay payload expansion.

**Acceptance:** Taskpane reports the same status semantics as PC; screenshots remain bounded; package/master/chart/table/XML fixes require confirmation; Office tests/typecheck/build pass.

**Commit:** `feat(office): close PowerPoint visual verification loop`.

## Task 7 — Cross-host golden evaluation and rollout controls

**Files:** shared fixtures/eval runner, feature flags, localization, docs.

Create adapter-independent golden tasks and expected receipts, including the reported cross-page consistency case. Add flags for contract generation, each host verifier, and autocorrection. Add all supported locale strings and privacy-safe diagnostics.

**Acceptance:** both adapters produce equivalent terminal status/check accounting; false-success and unintended-change regressions are zero in the fixture set; all locale keys exist; disabling flags restores prior behavior.

**Commit:** `test(presentation): gate verified agent editing`.

## Task 8 — Independent review and release verification

Review shared contracts, Agent Core state machine, PC implementation, Office implementation, and cross-host semantics independently. Resolve all Critical/Important findings with TDD. Run focused and full workspace tests, related typechecks, PC and Taskpane production builds, lint, format, diff checks, and package smoke tests.

Release both PC and Taskpane from the same verified main revision. Relay, WisWork PC pairing protocol, and Office manifest remain unchanged unless a reviewed implementation dependency explicitly requires otherwise.

## Rollback

Disable host verification/autocorrection flags independently. The shared contracts are additive and no persisted document migration is introduced. Reverting host adapters returns the existing generic QC and proposal/transaction flow; receipts stored in chat remain inert bounded metadata.
