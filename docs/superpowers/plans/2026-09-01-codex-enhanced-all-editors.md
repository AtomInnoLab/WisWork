# Codex Enhanced for all WisWork editors implementation plan

## Goal and non-goals

Implement the approved design in
`docs/superpowers/specs/2026-09-01-codex-enhanced-all-editors-design.md` on branch
`codex/codex-enhanced-all-editors`, based on `origin/main` commit `3c78aa0`.

Restore Codex as an optional, user-downloaded runtime for PC LaTeX, Slides, Docs, Sheets, and paired
Office Word, Excel, and PowerPoint. Standard Agent remains the default. Enhanced mode reuses WisWork
identity and WisUsage, activates only after restart, and may execute document mutations only through
the existing host transaction/proposal boundaries. The sole new elevated capability is raw Office
JavaScript/OOXML, with fresh explicit confirmation for every proposal.

Do not add OpenAI login or keys; generic shell, arbitrary filesystem, Git, browser control, free
network, direct writes, hot switching, silent fallback, Linux support, or bundled Codex binaries.
Do not assume the unmerged Office Taskpane `0.1.1`/manifest `0.3.4.0` branch is present; resolve that
version independently when this branch is integrated.

## Architecture

Shell owns one verified local Codex app-server and creates isolated logical document sessions. A
restored `@wiswork/codex-bridge` provides the pinned protocol/process/local WisUsage bridge and a
strict host-tool router; a restored `@wiswork/agent-runtime` selects Standard or Enhanced only at
startup while preserving Agent Core/Harness as the conversation, confirmation, receipt, and
terminal-truth authority. PC editors register their existing semantic skills and transaction
executors; Office sends Enhanced turns through the existing paired PC bridge while Office.js and
OOXML execution remains inside the Taskpane.

## Global constraints

- Runtime mode is `standard | enhanced`; requested and active modes are distinct and switching is
  restart-only.
- Component version, platform, size, SHA-256, archive layout, executable, and publisher policy are
  compile-time pinned. WisWork CDN and official OpenAI Release must resolve to identical bytes.
- Supported platforms are macOS arm64, macOS x64, and Windows x64. Unsupported platforms never
  download or launch.
- Codex receives no ambient credentials, `~/.codex`, generic shell, filesystem, Git, browser, free
  network, arbitrary MCP, direct writer, or cross-document capability.
- All parser, protocol, tool, prompt, output, screenshot, process, download, archive, and session
  inputs are strictly bounded before expensive work.
- Unknown tools, protocol variants, capability expansion, stale authority, and kill-switch state
  fail closed before dispatch.
- Mutation confirmation, transaction, history, rollback, post-readback, rendering, and completion
  receipt semantics remain host authoritative.
- Raw Office JavaScript/OOXML is a separate elevated tool, never a semantic change to the existing
  declarative `execute_office_js` tools.
- All seven hosts pass the compatibility gate before general availability. Per-host switches exist
  only for emergency rollback.
- Each task below ends in a scoped commit and an independent specification/code review before the
  next dependent task proceeds.

## Files and ownership map

- `packages/codex-bridge/**`: protocol, process, loopback WisUsage bridge, component manager, tool
  routing, generated pinned schema, limits, and security tests.
- `packages/agent-runtime/**`: startup-selected Standard/Enhanced runtime facade and lifecycle.
- `packages/agent-core/**`, `packages/agent-harness/**`: only shared hooks/events needed to host the
  alternative runtime; no duplicate conversation loop.
- `apps/shell/src/main/{codex-runtime,codex-ipc,enhanced-mode-component}.ts`: Shell lifecycle and
  renderer ownership.
- `apps/shell/src/{shared,preload,renderer}/**`: settings/download/restart/status UI and typed IPC.
- `apps/{latex,slides,docs,sheets}/src/**/ai/**`: host registration and current transaction adapters.
- `packages/office-bridge/**`, `apps/shell/src/main/office-*.ts`,
  `apps/office-addin/src/{pc-bridge,relay,agent}/**`: paired Enhanced capability and transport.
- `apps/office-addin/src/skills/{word,excel,powerpoint}/**`: elevated raw proposal compilers and host
  readback.
- `tools/codex/**`, root scripts, release workflows, notices, and docs: pinned distribution,
  packaging absence, evals, and rollback.

## Task 1: Replace removal regression with positive runtime policy contracts

**Files**

- Modify `tools/removed-optional-runtime.test.mjs` into
  `tools/optional-runtime-policy.test.mjs`.
- Modify root `package.json` and `package-lock.json`.
- Create `packages/agent-runtime/package.json`, `tsconfig.json`, `vitest.config.ts`,
  `src/contracts.ts`, `src/index.ts`, and `tests/contracts.test.ts`.
- Create `packages/codex-bridge/package.json`, `tsconfig.json`, `vitest.config.ts`, `src/types.ts`,
  `src/security.ts`, `src/index.ts`, and `tests/security-policy.test.ts`.

**Interfaces**

- `AgentRuntimeMode = 'standard' | 'enhanced'`.
- `EnhancedHost = latex | slides | docs | sheets | office-word | office-excel | office-powerpoint`.
- Strict `EnhancedRolloutPolicy`, capability declaration, component status, and safe error parsers.
- Denied-capability constants shared by launch and tests.

**Sequence**

1. Add RED policy tests proving the current removal test cannot represent optional runtime presence.
2. Add bounded strict contracts and positive tests for Standard defaults, all seven hosts, independent
   emergency switches, raw Office switch, unknown fields, accessors, oversized values, and denied
   capabilities.
3. Replace the blanket name/path prohibition with tests that require optionality, no bundled binary,
   no forbidden capabilities, and no active runtime when Standard is selected.
4. Add both packages to root test/typecheck chains and regenerate the lockfile.

**Acceptance**

- RED shows missing contracts/policy; GREEN package and policy tests pass.
- The new regression test would fail if an executable is bundled or a forbidden capability is
  exposed, but permits reviewed Enhanced source/UI names.
- Standard is the only default and all seven host switches are typed.

**Commit:** `feat(agent): define optional enhanced runtime policy`

## Task 2: Restore and harden the pinned app-server protocol

**Files**

- Restore and adapt `packages/codex-bridge/src/{json-rpc,app-server-client,process-manager}.ts` from
  the pre-removal implementation.
- Restore and adapt `src/generated/**`, `scripts/check-codex-schema.mjs`, and focused tests for JSON
  RPC, process lifecycle, requests, streams, and real app-server compatibility.
- Modify `packages/codex-bridge/src/index.ts` exports and package scripts.

**Interfaces**

- `CodexProcessManager` with fixed executable/config, minimal environment, deadlines, crash cleanup,
  and no inherited credentials.
- Version-bound `CodexAppServerClient` supporting initialize, thread/turn start, interrupt, events,
  and shutdown.
- A bound Responses-to-WisUsage Messages converter using only the approved one-call host-tool carrier.

**Sequence**

1. Restore historical adversarial tests and record RED against the skeleton.
2. Regenerate bindings for pinned `codex app-server` 0.147.0 and check its schema digest.
3. Restore bounded JSON-RPC and request/stream conversion without generic Responses tools.
4. Restore process lifecycle with shell/unified-exec/multi-agent/update-plan disabled and a minimal
   allowlisted environment.
5. Add explicit tests proving no filesystem, Git, browser, free-network, arbitrary MCP, or direct
   writer is advertised or callable.

**Acceptance**

- Historical protocol/hardening tests and a real pinned app-server contract pass.
- Unknown protocol/events fail with safe codes and no prompt or secret disclosure.
- Crash/stop rejects every pending request and leaves no process.

**Commit:** `feat(codex): restore bounded app-server protocol`

## Task 3: Implement three-platform on-demand component distribution

**Files**

- Restore and adapt `packages/codex-bridge/src/component-manager.ts` and tests.
- Create `tools/codex/manifest.json`, notices/licenses, and artifact fixtures.
- Restore and adapt `tools/install-enhanced-component.ts`, `tools/codex-release.test.mjs`,
  `tools/check-licenses.mjs`, and `tools/gen-third-party-notices.mjs`.
- Modify `.github/workflows/desktop-release.yml`, `.github/workflows/package-macos.yml`, and Windows
  release/signing steps where applicable.

**Interfaces**

- Strict component manifest for official 0.147.0 app-server-only assets on darwin-arm64,
  darwin-x64, and win32-x64.
- `EnhancedComponentManager` implementing primary CDN then official GitHub fallback with identical
  size/hash, safe extraction, signature policy, atomic promotion, update, launch recheck, and removal.

**Sequence**

1. Add RED manifest/platform/fallback/archive/signature tests.
2. Record official asset sizes, SHA-256 digests, exact entries, and executable paths; configure stable
   WisWork mirror URLs with the same digest.
3. Generalize the historical tar-only macOS manager to bounded tar.gz/zip extraction and platform
   publisher verification.
4. Add primary failure/fallback success, byte mismatch, redirect escape, traversal/link/device,
   duplicate/unexpected entry, interrupted install, TOCTOU, and concurrent-lock tests.
5. Add release checks proving no Codex executable is present in the base installer.

**Acceptance**

- All three platform fixtures select exactly one manifest record and reject substitution.
- A fallback is accepted only when its exact bytes match the pinned record.
- Package inspection proves on-demand delivery and notices/SBOM/license gates pass.

**Commit:** `build(codex): distribute verified app-server components`

## Task 4: Add the local WisUsage bridge and document-scoped tool router

**Files**

- Restore and adapt `packages/codex-bridge/src/{local-server,mcp-server,tool-router}.ts` and their
  hardening tests.
- Modify `packages/codex-bridge/src/index.ts`.
- Consume current `AgentSkill`, `AgentToolDef`, `AgentToolCall`, and `ToolExecution` from
  `@wiswork/agent-core`.

**Interfaces**

- `startResponsesBridge()` bound to numeric loopback, per-process credential, fixed WisUsage
  callback, concurrency/size/time bounds, and redacted diagnostics.
- `createDocumentToolSession()` with cryptographic credential, owner/document/generation identity,
  read-tool execution, mutation suspension, cancellation, and teardown.

**Sequence**

1. Restore RED server/router security tests, updating them for current Agent Core contracts.
2. Implement fixed local model bridge through injected authenticated WisUsage callback; do not expose
   tokens or arbitrary upstream URLs to Codex.
3. Implement read-only tool listing/calls and mutation forwarding into the existing suspended
   confirmation/transaction path.
4. Enforce exact session ownership, revision checks, unknown-tool rejection, tool/schema/input/output
   bounds, one-call carrier grammar, and cancellation.

**Acceptance**

- Invalid local credentials and cross-document calls never reach upstream or host tools.
- Mutations cannot execute without the host's existing confirmation/transaction authority.
- No registered tool provides any globally denied capability.

**Commit:** `feat(codex): route bounded host tools through WisUsage`

## Task 5: Restore Shell-owned runtime and restart-only settings

**Files**

- Create `apps/shell/src/main/{codex-runtime,codex-ipc,enhanced-mode-component}.ts`.
- Modify `apps/shell/src/main/{index,app-settings,tab-manager}.ts` only at reviewed integration points.
- Create `apps/shell/src/shared/{codex-api,enhanced-mode-api}.ts`; modify `home-api.ts`.
- Modify `apps/shell/src/preload/index.ts`, `renderer/src/{Home.tsx,enhanced-mode-view.ts,home.css}`,
  and localization strings/tests.
- Restore focused runtime, IPC, component, view, close, logout, and crash tests.

**Interfaces**

- Persisted `requestedAgentRuntime`; immutable-per-process `activeAgentRuntime`.
- Shell-owned runtime service exposing component status, download, update, remove, requested-mode
  change, restart requirement, host policy, and document-session IPC.

**Sequence**

1. Characterize current Standard, auth, Office pairing, tab-close, and settings behavior.
2. Add strict IPC/status APIs without activating Enhanced.
3. Wire component download/update/remove and safe localized UI.
4. Start Enhanced only during Shell startup when selected, signed in, installed, verified, and
   allowed; otherwise start no Enhanced session and surface a stable recovery state.
5. Wire logout, quit, document close, crash, and component update/removal lifecycle.

**Acceptance**

- Standard remains behaviorally unchanged and default.
- Selection changes set restart-required without hot switching or replay.
- Logout/quit/document close leave no process, local server, tool session, or Office grant.
- Unsupported platforms and kill switches fail closed without download.

**Commit:** `feat(shell): restore restart-bound Codex Enhanced runtime`

## Task 6: Provide one Agent Harness runtime adapter

**Files**

- Implement `packages/agent-runtime/src/{standard,enhanced}.ts` and tests.
- Modify only necessary extension points in `packages/agent-core/src/**` and
  `packages/agent-harness/src/**` with characterization tests.
- Add shared renderer runtime-client utilities under `packages/agent-runtime/src/renderer.ts`.

**Interfaces**

- `AgentRuntime.createSession({ host, document, skill, presentation?, events })`.
- Standard adapter around current `createAgentHarness` behavior.
- Enhanced adapter that converts app-server events into the same text/tool/plan/clarify/correction/
  confirmation/receipt events and uses the same host `AgentSkill`.

**Sequence**

1. Capture RED parity tests for ordinary text, read tool, suspended mutation, cancellation, reset,
   session replacement, provider error, plan, clarification, and verified receipt.
2. Implement Standard wrapper with zero behavior change.
3. Implement Enhanced session over Shell IPC and document tool session.
4. Enforce no silent fallback, no cross-runtime history, and no duplicate tool dispatch after crash.

**Acceptance**

- Both runtimes drive the same host skill and UI event contract.
- Standard tests remain unchanged; Enhanced failure never starts a Standard turn.
- Cancellation and reset preserve receipt/provider-pairing truth.

**Commit:** `feat(agent): select Standard or Enhanced runtime per restart`

## Task 7: Integrate PC LaTeX

**Files**

- Modify `apps/latex/src/renderer/ai/{AiPanel,agent-panel-session,latex-skill,transport}.ts(x)`.
- Create `apps/latex/src/renderer/ai/codex-tool-session.ts` only if the shared runtime adapter cannot
  directly express a LaTeX-specific authority operation.
- Modify bounded IPC in `apps/latex/src/{shared/ipc.ts,preload/index.ts,main/ipc.ts}`.
- Add runtime-mode and end-to-end workflow tests.

**Sequence**

1. Add RED mode-parity and complete read/propose/confirm/apply/compile/undo tests.
2. Register the current LaTeX skill through `AgentRuntime`; do not copy proposal logic.
3. Map document/revision/snapshot authority to the existing proposal workflow.
4. Add planning/clarification, restart, crash, cancellation, denial, concurrent-edit, and host
   kill-switch coverage.

**Acceptance**

- Enhanced uses existing proposal review, snapshot, compiler, verification, and undo.
- No mutation occurs through arbitrary project filesystem access.
- Standard LaTeX flow and dirty-close behavior remain unchanged.

**Commit:** `feat(latex): enable verified Codex Enhanced sessions`

## Task 8: Integrate PC Slides

**Files**

- Modify `apps/slides/src/renderer/ai/{AiPanel,agent-controller,transport,slides-skill}.ts(x)` and
  reuse `task-controller.ts`/`task-review.ts` unchanged unless a typed adapter is required.
- Modify bounded IPC in `apps/slides/src/{shared/ipc.ts,preload/index.ts,main/ai-ipc.ts}`.
- Add Enhanced runtime tests and extend the existing authority-bound Electron golden.

**Sequence**

1. Add RED tests for mode selection and a multi-slide canonical transaction through Enhanced.
2. Register the current Slides skill and production transaction executor through `AgentRuntime`.
3. Prove real mutation receipts, deterministic authority, Konva captures, strict visual reviewer,
   correction bounds, history rollback, and localized terminal receipts still compose.
4. Cover no-op, stale target, cancellation pre/post dispatch, session replacement, and kill switch.

**Acceptance**

- Enhanced reaches only current semantic tools; unsupported operations remain fail closed.
- Production transaction/verification lineage is unchanged and no direct PPTX write path appears.
- Standard Slides and the existing E2E artifact isolation gate pass.

**Commit:** `feat(slides): enable verified Codex Enhanced sessions`

## Task 9: Integrate PC Docs

**Files**

- Modify `apps/docs/src/renderer/ai/{AiPanel,agent-controller,transport,docs-skill}.ts(x)`.
- Modify bounded Docs main/shared/preload AI IPC at their existing registration points.
- Add Enhanced runtime, transaction, revision, and rollback tests.

**Sequence**

1. Add RED parity tests for bounded read and transactional text/format edits.
2. Register current Docs skill through `AgentRuntime` and current document command/atomic write path.
3. Add planning/clarification, no-op, protected/unsupported content, stale revision, cancellation,
   session replacement, verification, history, and kill-switch cases.

**Acceptance**

- Enhanced never edits DOCX/XML or filesystem directly.
- Existing protection, atomic write, pagination/render isolation, and undo semantics remain intact.
- Standard Docs behavior and builds pass.

**Commit:** `feat(docs): enable verified Codex Enhanced sessions`

## Task 10: Integrate PC Sheets

**Files**

- Modify `apps/sheets/src/renderer/ai/{AiChatPanel,agent-controller,transport,workbook-skill}.ts(x)`.
- Modify bounded Sheets main/shared/preload AI IPC at their current registration points.
- Add Enhanced runtime, workbook transaction, lazy-state, and rollback tests.

**Sequence**

1. Add RED parity tests for bounded reads and transactional value/formula/format operations.
2. Register current workbook skill through `AgentRuntime` and the existing workbook transaction path.
3. Preserve lazy sidecar bounds, loaded/live precedence, identity/revision guards, no-op behavior,
   cancellation, session replacement, verification, and kill switch.

**Acceptance**

- Enhanced cannot bypass workbook transactions or write XLSX/sidecars directly.
- Large/lazy workbooks remain bounded and stale callbacks cannot target a replacement workbook.
- Standard Sheets tests, native compatibility, and production build pass.

**Commit:** `feat(sheets): enable verified Codex Enhanced sessions`

## Task 11: Extend paired Enhanced mode to Office Taskpane

**Files**

- Modify `packages/office-bridge/src/index.ts` and tests.
- Modify Shell pairing/bridge/relay files under `apps/shell/src/main/office-*.ts` and typed APIs.
- Modify `apps/office-addin/src/{pc-bridge,relay}/**`, `agent/{transport,use-office-agent,host-runtime}.ts`,
  and connection/status UI/tests.
- Modify `services/wiswork-relay` only if the current session feature envelope cannot carry the
  non-callable Enhanced state without protocol expansion.

**Interfaces**

- Pairing/session statement containing active runtime mode, runtime instance, compatible component
  version, allowed host, raw-Office policy, expiry, and policy generation; no process handle or token.
- Office Enhanced transport over the current authenticated paired PC message path.

**Sequence**

1. Add RED pairing tests for absent Standard PC, Enhanced PC, stale restart, logout, host kill switch,
   component mismatch, expiry, and session replacement.
2. Extend existing pairing negotiation with a bounded non-callable feature statement; preserve exact
   legacy frames and persistent pairing behavior.
3. Select Enhanced transport only at Office Agent session creation; prohibit hot switching and
   replay.
4. Register current Word/Excel/PowerPoint semantic skills unchanged through the paired runtime.

**Acceptance**

- Office cannot use Enhanced unless the current paired PC proves an active compatible session.
- Taskpane receives no WisWork credential, app-server credential, runtime path, or process authority.
- Standard paired Office behavior and Relay compatibility remain unchanged.

**Commit:** `feat(office): pair Taskpane with Codex Enhanced runtime`

## Task 12: Add elevated raw Office JavaScript/OOXML proposals

**Files**

- Create `apps/office-addin/src/skills/shared/elevated-office-program.ts` and strict parser/tests.
- Create host adapters/tests under `skills/{word,excel,powerpoint}/elevated-*.ts`.
- Modify `agent/{host-runtime,use-office-agent,proposal-controller}.ts` and host skill registration.
- Reuse `skills/shared/office-write-transaction.ts` and existing readback/verification helpers.

**Interfaces**

- A distinct `propose_raw_office_edit` tool with discriminated Word/Excel/PowerPoint programs.
- Strict AST/OOXML package patch descriptors with compiled byte/node/statement/call/target/time/output
  maxima and no general JavaScript evaluator.
- Elevated proposal receipt binding host/document/session/revision/targets/program digest, confirmation,
  history snapshot, readback checks, rendering facts, and rollback.

**Sequence**

1. Add RED parser tests for safe host calls and every forbidden capability: network, dynamic load,
   eval/Function, storage, credentials, clipboard, cross-document/global access, unbounded control
   flow, excessive payload/calls/targets, malformed XML, external relationships, and scope expansion.
2. Implement a closed declarative AST and OOXML patch grammar; do not evaluate model-authored source
   with the JavaScript engine.
3. Add per-host compilers into exact Office.js allowlists and package-patch plans.
4. Route every proposal through structured preview and fresh explicit confirmation, then snapshot,
   execute, convergent readback/render proof, receipt, and rollback.
5. Reject raw operations in automatic correction; emit a new confirmation request instead.
6. Cover cancel before/after dispatch, timeout, uncertain completion, session replacement, changed
   proposal bytes, changed target authority, and no-retry semantics.

**Acceptance**

- No raw proposal can access network, credentials, another document, or an unapproved target.
- Confirmation is exact, fresh, single-use, and invalidated by any relevant drift.
- Proof unavailable after possible application returns `applied_unverified` without re-execution.
- Word, Excel, and PowerPoint each pass a real safe proposal/readback/rollback golden.

**Commit:** `feat(office): add confirmed raw JavaScript and OOXML edits`

## Task 13: Complete seven-host rollout, localization, telemetry, and release gates

**Files**

- Modify shared i18n catalogs and each host UI integration.
- Add safe aggregate Enhanced telemetry contracts and emitters.
- Add cross-host golden fixtures/tests and platform/component E2E drivers.
- Modify `.github/workflows/ci.yml`, desktop release workflows, release documentation, notices, and
  rollback runbooks.

**Sequence**

1. Add all 19 locale keys for download, restart, unavailable, plan, clarify, confirmation,
   correction, verified, applied-unverified, rollback, and removal states; wire real production UI.
2. Add closed-enum, fail-open telemetry for component phases and per-host plan/dispatch/verify/
   complete outcomes with no IDs or content.
3. Add one executable seven-host golden outcome catalog using production host adapters, transaction/
   proposal paths, deterministic checks, and render review where available.
4. Add macOS arm64/x64 and Windows x64 clean-checkout install/download/restart/launch/update/remove,
   offline, corrupt, fallback, signature, crash, and kill-switch jobs.
5. Add release rollout/incident instructions and prove each host/raw switch restores Standard behavior.

**Acceptance**

- Seven hosts are consistently available only when the entire release gate passes.
- All visible states are localized and receipt-derived; telemetry cannot change task semantics.
- Base packages contain no Codex executable; downloaded component identity is verified before every
  launch.
- Emergency global, per-host, component-version, and raw-Office rollback are independently proven.

**Commit:** `test(codex): gate seven-host Enhanced rollout`

## Final verification and integration

1. Run every restored/new package test and typecheck.
2. Run full tests for Shell, LaTeX, Slides, Docs, Sheets, Office Taskpane, Agent Core/Harness,
   Office Bridge, Presentation Verification, and Relay.
3. Run root `npm test`, `npm run typecheck`, `npm run lint`, `npm run format:check`,
   `npm run build:all`, license/notices/SBOM checks, `git diff --check`, and production package
   inspection from a clean checkout. Use a writable Cargo target for native/Relay checks.
4. Execute Electron/browser E2E with compile-time-only test seams and verify production artifacts do
   not contain those seams.
5. Execute component contract tests with the pinned official app-server on all three platform jobs.
6. Perform an independent broad architecture/security review of the full branch, including source to
   mutation sinks for raw Office proposals, component supply chain, local bridge, pairing, and
   cancellation/session races.
7. Fix Critical/Important findings, rerun affected and broad gates, then use
   `finishing-a-development-branch` to prepare the PR or integration choice.

## Migration and rollback

No document migration is introduced. Existing settings without a runtime field parse as Standard.
Historical `enhanced` settings are accepted only through the new strict requested-mode parser and do
not activate until component verification and restart. Old cached components are never discovered
by path scanning; only the new signed manifest/version records are eligible.

Rollback proceeds in this order: raw Office switch, affected host switch, global Enhanced switch,
component-version allowlist removal, then code revert. Undispatched work is cancelled; applied work
retains existing history and receipt truth. Removing runtime code leaves cached component bytes inert
until a separately reviewed cleanup targets the exact app-private directory.

## Plan self-review

- Every approved host, platform, permission denial, raw Office constraint, download source, identity
  rule, restart rule, verification boundary, rollout switch, and rollback layer is assigned to a
  deliverable.
- Dependencies are ordered: contracts -> protocol -> component -> bridge/tools -> Shell -> runtime
  adapter -> PC hosts -> Office pairing -> elevated Office -> rollout.
- Shared aggregation files are touched serially to avoid concurrent package-lock, Shell index, Office
  host runtime, Relay session, and workflow conflicts.
- There are no unresolved placeholders. The pinned release is 0.147.0; Task 3 records authoritative
  sizes/hashes from official assets and requires the WisWork mirror to match them exactly.
- The plan does not revive the old single-platform pilot wholesale or weaken the post-removal safety
  regression; it converts that regression into a positive optional-runtime isolation gate.
