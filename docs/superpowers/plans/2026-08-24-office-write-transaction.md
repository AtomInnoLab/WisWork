# Office Write Transaction Implementation Plan

## Goal and architecture

Implement the approved three-state Office write contract without expanding authority. Each host keeps domain-specific snapshots and semantic verifiers; a small shared bounded convergence/error vocabulary avoids divergent retry behavior while host adapters retain recovery ownership.

## Global constraints

- No Manifest, Relay, PC, model, or permission changes.
- Never restore over a third/concurrent state.
- Never report `confirmed` without semantic proof.
- Never report a failed operation as unchanged without restoration proof.
- All retries are bounded, abortable, and deterministic in tests.

## Files

- `apps/office-addin/src/agent/proposal-controller.ts`: stable terminal failure vocabulary only.
- `apps/office-addin/src/agent/use-office-agent.ts`: safe user copy and no-same-turn retry policy.
- `apps/office-addin/src/diagnostics/office-diagnostics.ts`: local safe diagnostic codes.
- `apps/office-addin/src/skills/shared/office-write-transaction.ts`: bounded convergence primitives and state classification.
- Host adapters/skills under `skills/word`, `skills/excel`, and `skills/powerpoint`: snapshots, receipts, verification, recovery.
- Host-focused tests plus `agent-session`, `diagnostics`, and `relay-session` compatibility tests.

## Task 1 — Shared contract and Word

Deliver a shared bounded convergence helper and harden Word `write_document` plus declarative operations.

Acceptance:

- RED tests reproduce delayed readback, declarative partial commit, style drift, unexpected formatting, and concurrent edit before recovery.
- GREEN proves applied/restored/uncertain classification and no overwrite of third state.
- Existing Word complex writes remain green.
- Scoped commit with focused Word, diagnostics, and typecheck evidence.

## Task 2 — Excel

Resolve stable worksheet identity and add operation-specific snapshots, bounded verification, and compare-and-restore for direct Excel mutations.

Acceptance:

- RED tests reproduce worksheet reorder, sync rejection after partial cell/style/note mutation, delayed readback, and concurrent target edit.
- GREEN never writes the wrong worksheet, restores only attributable partial state, and reports uncertain for third state.
- Cell, range, structure, workbook, resize, chart/pivot, CSV, and image paths retain limits and confirmation.
- Scoped commit with focused Excel tests and typecheck evidence.

## Task 3 — PowerPoint

Extend transaction safety from package replacement to direct Office.js text/shape/duplicate/declarative operations.

Acceptance:

- RED tests reproduce partial declarative sync, delayed visibility for text/move/delete/duplicate, cancellation, and concurrent slide edits.
- GREEN uses semantic snapshots, bounded convergence, and safe compare-and-restore.
- Background-only package normalization remains strict outside `<p:bg>` and generic package verification is unchanged.
- Scoped commit with focused PowerPoint tests and typecheck evidence.

## Task 4 — Integration and release

Wire safe error copy/diagnostics, run full verification, and perform independent broad review.

Acceptance:

- New failure codes cannot tear down Relay v2 sessions and do not expose document data.
- Full Office tests, typecheck, lint, format, theme, production build, and diff checks pass.
- Independent reviewers report no Critical or Important findings after at most two fix rounds.
- Release handoff states Taskpane-only deployment and the required real-host acceptance matrix.

## Security and rollback

Snapshots remain bounded and in-memory; no document body is logged or sent to Relay. Recovery uses stable IDs and compare-before-restore. Roll back by redeploying the prior Taskpane build.
