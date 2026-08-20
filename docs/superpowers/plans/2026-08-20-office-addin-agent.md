# Office Add-in Agent MVP Implementation Plan

**Specification:** `docs/superpowers/specs/2026-08-20-office-addin-agent-design.md`

## Goal and non-goals

Deliver a fail-closed Office task-pane Agent MVP that uses WisWork browser OAuth and the fixed WisUsage messages service, streams through `@wiswork/agent-core`, reads the active selection, and requires explicit confirmation before replace or append operations.

Research retrieval, citations, persistent sessions, bulk edits, formatting, files, connectors, Outlook, and production Gateway enablement are not part of this implementation.

## Architecture and constraints

The Office renderer owns only browser PKCE state, an in-memory token session, AgentLoop presentation state, and bounded Office.js tools. It calls fixed, validated WisWork endpoints; endpoint/model/region/auth values cannot be supplied by model output or user prompts. All mutations are two-phase: a tool creates a proposal, then a separate user confirmation revalidates selection state and performs Office.js mutation.

Global constraints:

- No token, authorization code, PKCE verifier, refresh token, or upstream response body in logs, persisted UI state, document content, or user-visible errors.
- Missing or invalid Gateway configuration fails closed.
- OAuth uses PKCE S256, exact state matching, single callback parameters, expected issuer checks, and one refresh retry.
- Access/refresh tokens remain in memory; PKCE state is session-only and one-time.
- Selection/proposal/tool payloads are length-capped and schema validated.
- Every runtime behavior starts with a failing test and records RED/GREEN evidence.
- Real Gateway acceptance remains blocked until callback registration, PKCE/CORS, refresh, and safe-error contracts are confirmed externally.

## Files and responsibilities

- `apps/office-addin/src/config.ts`: validate non-secret browser configuration and fixed defaults.
- `apps/office-addin/src/auth/pkce.ts`: cryptographic verifier, state, and S256 challenge helpers.
- `apps/office-addin/src/auth/browser-auth.ts`: authorization redirect, callback validation/exchange, in-memory session, refresh once, logout.
- `apps/office-addin/src/agent/transport.ts`: WisUsage-compatible SSE transport implementing `AgentTransport`.
- `apps/office-addin/src/agent/office-skill.ts`: read/propose tools with strict input bounds.
- `apps/office-addin/src/agent/proposal-controller.ts`: pending proposal lifecycle, stale-selection check, confirmed mutation.
- `apps/office-addin/src/agent/use-office-agent.ts`: AgentLoop and React-facing chat/activity state.
- `apps/office-addin/src/App.tsx`, `styles.css`: auth, chat, tool activity, proposal preview, confirm/reject/stop/logout UI.
- `apps/office-addin/src/office-document.ts`: append capability and selection fingerprint support.
- `apps/office-addin/src/taskpane.html`, `public/manifest.xml`, `vite.config.ts`: callback route and explicit allowed origins.
- `apps/office-addin/tests/*`: unit and integration-level tests for each boundary.
- `apps/office-addin/README.md`: configuration, Gateway prerequisite, local/manual acceptance steps.
- `apps/office-addin/package.json`, root lockfile: workspace dependencies only where required.

## Task 1: Fail-closed browser OAuth session

**Produces:** validated runtime configuration and a browser OAuth client that can start PKCE login, consume one callback, hold tokens in memory, refresh once, and clear all session material.

**Files:** `src/config.ts`, `src/auth/pkce.ts`, `src/auth/browser-auth.ts`, `tests/config.test.ts`, `tests/pkce.test.ts`, `tests/browser-auth.test.ts`, Vite environment types/configuration.

**Acceptance criteria:**

- Invalid/missing URLs, client ID, issuer, or non-HTTPS production callback yield a stable unavailable state.
- Authorization URL contains state, S256 challenge, exact redirect URI, and configured client ID.
- Callback rejects duplicate/missing code/state, wrong state, wrong issuer, and replay.
- Exchange request sends verifier but no secrets are persisted or surfaced.
- Authenticated fetch refreshes exactly once after a 401 and logs out after a second 401.
- Logout clears memory and session PKCE keys.

**Implementation sequence and verification:**

1. Add tests for configuration, PKCE, callback validation, replay, refresh, and safe errors.
2. Run `npm run test -w @wiswork/office-addin -- tests/config.test.ts tests/pkce.test.ts tests/browser-auth.test.ts`; RED must be missing modules/behavior.
3. Implement the smallest browser auth boundary and rerun to GREEN.
4. Run Office workspace typecheck and lint.
5. Scoped commit: `feat(office-addin): add fail-closed browser oauth`.

## Task 2: Bounded Office Agent tools and transport

**Produces:** fixed WisUsage streaming transport, Office AgentSkill, and two-phase proposal controller.

**Files:** `src/agent/transport.ts`, `src/agent/office-skill.ts`, `src/agent/proposal-controller.ts`, `src/office-document.ts`, `tests/transport.test.ts`, `tests/office-skill.test.ts`, `tests/proposal-controller.test.ts`, existing Office adapter tests.

**Acceptance criteria:**

- Transport ignores prompt/tool attempts to alter endpoint, model, region, or authorization.
- SSE text/tool calls/stop/error are normalized into AgentLoop callbacks; cancel aborts and completes once.
- Network/HTTP errors expose only stable stage/status, never response bodies.
- `read_selection` returns capped text.
- proposal tools reject extra/missing/oversized arguments and never mutate Office.
- confirm compares a selection fingerprint, rejects stale proposals, and calls replace/append only after explicit confirmation.
- reject/new turn/logout invalidates the proposal.

**Implementation sequence and verification:**

1. Add transport, skill, and proposal behavior tests; run targeted tests and record expected RED.
2. Implement transport parsing and strict tool/proposal boundaries; rerun to GREEN.
3. Add/adjust Office adapter tests for append and fingerprints; keep all Office tests GREEN.
4. Run Office workspace typecheck and lint.
5. Scoped commit: `feat(office-addin): add bounded agent tools and transport`.

## Task 3: AgentLoop UI, security configuration, and operator docs

**Produces:** usable signed-out/ready/working/proposal/error task-pane states connected to AgentLoop, plus security and manual acceptance documentation.

**Files:** `src/agent/use-office-agent.ts`, `src/App.tsx`, `src/styles.css`, `src/main.tsx`, `src/taskpane.html`, `public/manifest.xml`, `vite.config.ts`, `tests/agent-session.test.ts`, `tests/manifest.test.ts`, `README.md`, package metadata/lockfile.

**Acceptance criteria:**

- Unconfigured builds show unavailable and cannot send prompts.
- Signed-in users can stream text, stop a run, see tool activity, and start a new turn only after pending proposals are resolved/invalidated.
- Proposal UI shows before/after, Confirm, and Reject; no write occurs during rendering.
- Logout clears agent history and pending proposal.
- CSP/connect origins are generated from validated configuration and do not permit arbitrary model/tool URLs.
- Build includes callback/taskpane routes and manifest resources.
- README lists required Gateway registration and a manual acceptance checklist without suggesting tokens in environment variables or browser storage.

**Implementation sequence and verification:**

1. Add controller/session and manifest/CSP tests; run targeted tests and record RED.
2. Implement hook/UI states and deterministic callback routing; rerun to GREEN.
3. Update styles, manifest/build configuration, and operator documentation.
4. Run Office tests, typecheck, build, ESLint, format check, XML/CSP structural checks.
5. Scoped commit: `feat(office-addin): wire agent chat and confirmed edits`.

## Independent review and final verification

After each task, a fresh reviewer checks the task commit for specification compliance, security, tests, and code quality. Critical/Important findings return to the implementer for at most two fix-and-review rounds.

After all tasks:

1. Run a broad independent review across the complete Agent diff.
2. Run fresh Office tests/typecheck/build/lint/format/XML checks.
3. Run root typecheck and root tests. If the existing Sheets release build again exhausts shared temporary storage, record the exact environmental failure and separately run all remaining app suites; do not describe the full suite as passing.
4. Search the diff/build output for token/verifier/code logging or persistence patterns.
5. Record manual acceptance as blocked until Gateway callback/PKCE/CORS/refresh contracts are confirmed and a compatible environment is available.

## Rollback, migration, and release

No migration or persistent user state is introduced. Rollback removes the Agent/auth additions and leaves the selection-only Office lab intact. Production release is gated on Gateway callback registration, verifier validation, one-time codes, allowed CORS origins, refresh behavior, CSP origins, and packaged manual acceptance; development builds without valid configuration remain fail-closed.
