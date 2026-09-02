# Codex Enhanced for all WisWork editors

## Status

Approved product design for implementation planning. This specification restores Codex as an
optional WisWork runtime and extends it consistently to PC LaTeX, Slides, Docs, Sheets, and the
Office Taskpane hosts for Word, Excel, and PowerPoint.

## Product decision

WisWork keeps the existing Standard Agent as the default. A signed-in user may explicitly download
and enable **Codex Enhanced** on a supported WisWork PC. Enhanced mode becomes active only after an
application restart. Office Taskpane hosts may use it only while paired with an Enhanced-enabled PC.

The first release is all-on across the seven editor hosts. It is not acceptable to advertise Codex
Enhanced in only a subset of those hosts. Each host nevertheless has an independent server-side
emergency kill switch so an unsafe host integration can be disabled without withdrawing the whole
PC release.

## Goals

- Restore the previously removed local Codex app-server integration without weakening Standard
  Agent behavior.
- Support one consistent Enhanced-mode lifecycle in PC LaTeX, Slides, Docs, Sheets, and Office Word,
  Excel, and PowerPoint.
- Reuse the user's existing WisWork identity, membership, credits, and WisUsage provider path.
- Download a fixed, verified Codex runtime on demand instead of increasing the base installer size.
- Preserve the existing editor-specific transaction, proposal, confirmation, verification,
  history, rollback, and completion-receipt boundaries.
- Permit raw Office JavaScript/OOXML only through the existing elevated proposal flow with explicit
  confirmation and post-write verification.
- Make planning and clarification visible before mutation when a request is ambiguous, high risk,
  or materially broader than the captured scope.

## Non-goals

- Codex Enhanced is not the default and is never downloaded or enabled silently.
- No OpenAI or ChatGPT login is introduced.
- WisWork does not read `~/.codex`, OpenAI API keys, ChatGPT credentials, model environment
  variables, or another Codex installation's state.
- The runtime receives no generic shell, arbitrary filesystem, Git, browser-control, or unrestricted
  network capability.
- The runtime cannot write an editor document outside the existing transaction/proposal layer.
- Linux is not supported in the first release.
- Hot switching between Standard and Enhanced mode is not supported.
- A failed Enhanced run is not silently replayed through Standard Agent or vice versa.

## Supported hosts and platforms

The first generally available Enhanced release is enabled as one compatibility set:

| Surface         | Hosts                       | Runtime location                     |
| --------------- | --------------------------- | ------------------------------------ |
| WisWork PC      | LaTeX, Slides, Docs, Sheets | Local Codex app-server child process |
| Office Taskpane | Word, Excel, PowerPoint     | Paired WisWork PC Codex app-server   |

Supported desktop platforms are:

- macOS arm64
- macOS x64
- Windows x64

Linux and unsupported architectures show Enhanced mode as unavailable and do not download or launch
anything.

## User experience

### Discovery and download

The account/settings menu shows Standard Agent and Codex Enhanced. Standard remains selected by
default. Selecting Enhanced for the first time opens a bounded download panel showing:

- the fixed Codex runtime version;
- platform and architecture;
- download size;
- publisher and license notice;
- WisWork CDN as the primary source and the official OpenAI GitHub Release as fallback;
- a statement that the component is optional and can be removed;
- a **Download and enable** action.

The user downloads the runtime from inside WisWork PC. There is no separate user-facing installer.
The settings panel also exposes the exact verified artifact URL for diagnostics and manual
enterprise mirroring. The release manifest uses these stable URL shapes:

- Primary: `https://office.8-216-134-194.sslip.io/components/codex/<version>/<artifact>`
- Fallback: `https://github.com/openai/codex/releases/download/rust-v<version>/<artifact>`

Both sources must serve bytes matching the same release-pinned size and SHA-256 digest. Redirects
outside the approved host allowlist fail closed.

### Activation and switching

After a successful download and verification, the UI says that restart is required. The selected
mode is persisted, but the running process does not switch. On restart, WisWork validates the
component again before launching it. A runtime launch or handshake failure leaves the current task
unstarted and presents a safe recovery choice; it never replays the task in Standard mode.

Switching back to Standard, changing Codex versions, removing the component, or re-enabling after a
kill-switch transition also takes effect only after restart.

### Planning, clarification, and confirmation

Enhanced mode follows the same observable workflow in all seven hosts:

1. Capture the authoritative editor/document/session scope.
2. If the requested result or target is ambiguous, ask a concise clarification before dispatch.
3. For a multi-step or risky task, show a bounded localized plan describing affected scope,
   operations, verification, and whether confirmation is required.
4. Freeze the approved scope and acceptance contract.
5. Dispatch only operations covered by that contract.
6. Verify deterministic postconditions and, when required, rendered output.
7. Produce a receipt-derived terminal result rather than an unsupported model claim.

A plan is informational unless the operation is already governed by a confirmation boundary. Raw
Office JavaScript/OOXML always requires an explicit per-proposal confirmation regardless of whether
a plan was shown.

### Office experience

The Office Taskpane shows Enhanced only after pairing with WisWork PC and receiving a capability
statement that the PC is signed in, Enhanced-enabled, compatible, and not kill-switched for that
host. Pairing never transfers the Codex binary, process handle, WisWork token, or model credential to
Office. Disconnect, PC logout, PC restart, document replacement, or capability expiry cancels the
active run and requires re-pairing or revalidation.

## Architecture

### Component layout

Restore the previous architecture as reviewed, shared packages rather than editor-specific process
launchers:

- Shell owns component download, verification, installation, selection, launch, restart lifecycle,
  health, update, removal, and diagnostics.
- A restored runtime package owns the bounded Codex app-server process and protocol lifecycle.
- A restored bridge package maps Agent Harness turns and host capabilities to the Codex app-server
  protocol without exposing editor internals.
- Agent Core and Agent Harness remain the single conversation, cancellation, confirmation, receipt,
  and terminal-response authority.
- Each PC editor supplies an adapter containing only its already approved semantic operations and
  transaction executors.
- The existing Office loopback bridge carries bounded Codex requests between Taskpane and the paired
  PC. Office.js execution remains in the Taskpane; the PC never mutates the Office document.

### Runtime ownership

There is at most one Shell-owned Codex app-server process per signed-in desktop session. Individual
editor runs use isolated logical sessions with document, host, and generation identities. A run
cannot access another editor, document, conversation, or prior session. Closing a document cancels
its pending work and releases its session. Logout stops the process and revokes all Office bridge
capabilities.

The runtime executable is launched by fixed absolute path with a fixed argument allowlist and a
minimal environment. Working directories, inherited handles, search paths, proxy variables, and
credential-related environment variables are removed unless explicitly required and reviewed.

### Request flow

For a PC editor:

`Editor UI -> Agent Harness -> Codex bridge -> local app-server -> bounded tool request -> editor
transaction/proposal executor -> verification -> receipt -> Agent Harness`

For Office:

`Taskpane Agent Harness -> paired loopback capability -> PC Codex bridge/app-server -> Taskpane tool
proposal -> user confirmation -> Office.js executor -> readback/visual verification -> receipt`

The model may propose work, but only the host adapter can authorize or execute it. Tool output is
bounded, detached, schema-validated, and privacy-filtered before returning to the runtime.

## Capability policy

### Allowed capabilities

Codex Enhanced receives only:

- bounded conversation context supplied by Agent Harness;
- host-specific semantic read tools;
- host-specific semantic mutation proposals that execute through the existing transaction layer;
- bounded render/screenshot facts used by the existing verification loop;
- raw Office JavaScript/OOXML proposals under the elevated policy below.

### Explicitly denied capabilities

The runtime receives no:

- generic shell or process execution tool;
- arbitrary filesystem read or write;
- Git access;
- browser automation or browser-control tool;
- unrestricted network access;
- direct document write primitive that bypasses transactions, proposals, history, or verification.

These are architectural denials, not prompt instructions. The app-server is started without MCP
servers or host adapters that could provide them. Unknown tool names and capability expansion fail
closed before dispatch.

### Host semantic tools

LaTeX, Slides, Docs, and Sheets expose the same semantic read/write families currently available to
Standard Agent, subject to their existing bounds and transaction semantics. Codex integration must
not advertise operations that the production host cannot execute and verify.

Office Word, Excel, and PowerPoint expose the existing bounded Office semantic tools plus the
elevated raw Office capability. Ordinary semantic mutations retain their existing preview,
confirmation, and verification rules.

## Elevated raw Office JavaScript/OOXML

Raw Office JavaScript and OOXML are the only additional powerful capability approved for Codex
Enhanced.

### Proposal boundary

- Every raw call is an elevated proposal and requires fresh, explicit user confirmation.
- Confirmation is scoped to the exact bounded code/package patch, host, document, target set, and
  operation digest shown in the preview.
- Approval is single use. Editing, regeneration, retry, correction, session replacement, or target
  drift requires another confirmation.
- A visual-review correction pass cannot invoke raw Office JavaScript/OOXML automatically. It must
  return to the user with a new proposal and confirmation request.

### Static restrictions

The parser and proposal compiler reject:

- network requests, sockets, WebSockets, or external resource loads;
- dynamic imports, script injection, `eval`, `Function`, or equivalent dynamic execution;
- access to cookies, local/session storage, clipboard, credentials, auth tokens, or environment;
- cross-document, cross-workbook, cross-presentation, add-in-global, or account-level access;
- unbounded loops, recursion, payloads, calls, targets, XML size, or output;
- host APIs outside the reviewed allowlist;
- writes whose target scope cannot be frozen before dispatch.

The implementation enforces explicit maximums for source bytes, AST nodes, statements, Office API
calls, target count, XML bytes, output bytes, execution time, and cancellation latency. Exact numeric
budgets belong in the implementation plan and shared contracts; they may only be reduced by remote
policy, never increased beyond compiled maxima.

### Execution and proof

- Capture document/session/revision authority and a package/history snapshot before dispatch.
- Execute inside the existing elevated Office proposal executor, not a new generic JavaScript
  runtime.
- Validate proposal identity again immediately before execution.
- Apply the existing timeout, cancellation, and session-replacement guards.
- Read back every deterministic postcondition and capture rendered evidence where required.
- Return `verified` only when all frozen checks pass.
- Return `applied_unverified` when mutation may have applied but proof is unavailable or stale.
- Preserve the snapshot/history identifier for rollback whenever the host supports it.
- Never silently retry a raw mutation after uncertain completion.

## Download, integrity, and installation

### Pinned artifacts

The component manifest is committed and release-reviewed. It contains one record per supported
platform with:

- Codex version;
- primary and fallback HTTPS URLs;
- exact compressed byte size;
- SHA-256 digest;
- expected archive entries and extracted byte bounds;
- executable relative path;
- publisher/signing requirements;
- license and source notices.

The initial version is selected during implementation from an official OpenAI release that exposes
all three required app-server artifacts. A version change is a reviewed manifest change and cannot
be supplied solely by a remote response.

### Verification pipeline

1. Resolve only an approved manifest entry for the local platform.
2. Download to an app-private temporary file with time, redirect, byte, and concurrency limits.
3. Verify compressed size and SHA-256 before extraction.
4. Inspect archive paths and reject absolute paths, traversal, links, devices, duplicates, or
   unexpected files.
5. Extract into a new versioned staging directory with entry and total-byte limits.
6. Verify the extracted manifest and executable signature/publisher policy.
7. Atomically promote staging to the immutable version directory.
8. Record only component version, platform, digest, install state, and safe failure code.
9. On launch, revalidate version, path containment, executable identity, and signature.

macOS requires the approved code-signing/notarization policy and Windows requires the approved
Authenticode publisher policy. A failed source, digest, extraction, signature, or launch check does
not fall back to unverified bytes.

### Updates and removal

Updates download beside the active immutable version and activate only after restart. A failed
update leaves the old verified version active. Removal stops new Enhanced sessions, takes effect
after restart if a run is active, and deletes only the resolved app-private component directory.

## Identity, provider, and privacy

WisWork authentication remains authoritative. The Codex bridge sends model requests through the
same WisUsage account, provider configuration, entitlement, quota, refresh, and logout behavior as
Standard Agent. It does not expose WisWork tokens to the app-server or Office.

Only the minimum bounded task context and tool facts cross into Codex. Raw document inventories,
filesystem paths, runtime IDs, credentials, full screenshots beyond the approved verifier budget,
and unrelated editor content are excluded. Logs and telemetry use closed safe-code enums and
aggregate counts; they contain no prompt text, document text, file paths, session tokens, element
IDs, proposal source, OOXML, screenshots, or receipt identifiers.

## State model and failure semantics

The Shell component state is one of:

- `not_installed`
- `downloading`
- `verifying`
- `installed_restart_required`
- `ready`
- `update_available`
- `removal_restart_required`
- `blocked_by_policy`
- `unsupported_platform`
- `failed_safe`

An editor run independently tracks captured authority, frozen plan/contract, dispatch state,
mutation evidence, verification evidence, and terminal receipt. Component state changes cannot
retarget or replay a run.

Failure behavior:

- Before dispatch: return unchanged/cancelled, clarification, or a stable unavailable state.
- After proved mutation: preserve applied truth and reconcile to verified, applied-unverified, or
  failed-with-rollback evidence as allowed by the existing receipt contract.
- On runtime crash: do not restart and replay the task. A later user action may launch a fresh
  session after authority recapture.
- On Office disconnect: cancel undispatched work; reconcile already-dispatched work using existing
  Office readback and abandoned-receipt rules.
- On Standard/Enhanced selection mismatch: require restart; do not mix histories or tool calls.

## Feature flags and rollout

Compiled support, release manifest availability, account entitlement, and server policy must all
allow Enhanced mode. The rollout uses:

- one global Enhanced availability flag;
- one component-version allowlist;
- independent host kill switches for LaTeX, Slides, Docs, Sheets, Word, Excel, and PowerPoint;
- an independent raw Office JavaScript/OOXML kill switch;
- existing planning, verified-completion, visual-review, and autocorrection flags.

All seven host adapters must pass acceptance before general availability. Host kill switches are
emergency rollback controls, not staged product exposure. Autocorrection remains off by default
unless separately approved; raw Office capability is never eligible for automatic correction.

## Delivery and ownership

- **Shell/runtime owner:** component manifest, download/install/launch lifecycle, restart UX,
  signatures, diagnostics, and removal.
- **Agent platform owner:** Codex bridge, Agent Harness integration, session isolation, capability
  negotiation, bounded protocol, receipts, and telemetry contracts.
- **PC editor owners:** LaTeX, Slides, Docs, and Sheets semantic adapters, transaction authority,
  readback, rendered verification, history, and rollback.
- **Office owner:** Taskpane pairing, Word/Excel/PowerPoint adapters, elevated raw proposal compiler,
  confirmation UI, Office.js execution, readback, and rollback evidence.
- **Release/security owner:** artifact mirroring, pinned hashes, license/source notices, signing and
  notarization gates, SBOM, malware scan, and kill-switch readiness.
- **QA owner:** cross-host golden corpus, platform matrix, offline/failure tests, security abuse
  cases, and manual release acceptance.

## Verification strategy

### Shared contract tests

- Capability negotiation rejects unknown, expanded, or host-incompatible tools.
- Standard mode contains no Codex launch or routing path.
- Enhanced mode cannot receive shell, filesystem, Git, browser, free-network, or direct-write tools.
- Runtime inputs, outputs, screenshots, tool calls, plans, confirmations, receipts, and telemetry are
  strictly parsed and bounded.
- Restart-only activation and no-silent-fallback behavior are deterministic.

### Component security tests

- Primary and fallback sources accept only identical pinned bytes.
- Redirect, truncation, overrun, digest mismatch, archive traversal, symlink, device, duplicate,
  unexpected entry, signature mismatch, TOCTOU, and interrupted promotion fail closed.
- Platform/architecture selection never substitutes another artifact.
- Launch environment contains no credential or model-provider secrets.
- Remove/update operations cannot escape the app-private component root.

### Editor acceptance

Each of the seven hosts must prove:

- Standard Agent remains unchanged when Enhanced is absent, disabled, blocked, or removed.
- Enhanced planning/clarification freezes the intended document and target scope before mutation.
- Supported semantic operations run through the production transaction/proposal executor.
- Deterministic and rendered postconditions produce receipt-derived terminal truth.
- Cancellation before dispatch writes nothing; cancellation after dispatch preserves applied truth.
- Session/document replacement cannot mutate or reconcile into the replacement run.
- No-op requests return unchanged without synthetic mutation evidence.
- Rollback/history evidence is preserved where supported.
- Host kill-switch rollback restores Standard behavior without migration.

### Raw Office acceptance

- Word, Excel, and PowerPoint each cover a safe bounded raw proposal, explicit confirmation,
  execution, readback, rendered proof where applicable, and rollback.
- Network, dynamic execution, credentials, storage, cross-document access, excessive code/XML,
  excessive calls/targets, timeout, malformed output, and scope expansion all fail before dispatch.
- Changed proposal bytes or authority after preview invalidate confirmation.
- Correction attempts return a new confirmation request and never auto-dispatch raw code.
- Uncertain completion never re-executes and reports `applied_unverified` with safe rollback evidence.

### Release gates

- Full workspace tests, typechecks, lint, formatting, license/notices, SBOM, and production builds
  pass on a clean checkout.
- macOS arm64, macOS x64, and Windows x64 packages pass install, download, restart, launch, update,
  removal, offline, corrupt-download, and kill-switch acceptance.
- Package inspection proves the base installer does not contain the optional Codex executable.
- Office pairing tests prove the Taskpane cannot use Enhanced without a current paired PC grant.
- A seven-host golden task suite reports equivalent bounded outcomes without claiming unsupported
  capabilities.

## Rollback

Rollback is layered and recoverable:

1. Disable raw Office JavaScript/OOXML independently.
2. Disable an affected host adapter with its host kill switch.
3. Disable Enhanced globally while leaving Standard Agent available.
4. Remove a bad component version from the signed allowlist; installed bytes remain inert and cannot
   launch.
5. Ship a PC revert that removes runtime/bridge UI and code paths while leaving existing component
   cache inert until a separately reviewed safe cleanup.

No rollback requires a document migration. In-flight undispatched work is cancelled. Already-applied
mutations retain their existing history, rollback, and receipt truth.

## Success criteria

- A user can download a fixed verified Codex runtime in WisWork PC, restart, and use Enhanced mode in
  LaTeX, Slides, Docs, Sheets, Word, Excel, and PowerPoint.
- The same WisWork account and WisUsage billing path are used with no OpenAI credential flow.
- Standard Agent remains the default and is behaviorally unchanged.
- Office Enhanced mode works only through an approved paired PC.
- The runtime has no generic shell, arbitrary filesystem, Git, browser, free-network, or direct-write
  capability.
- Raw Office JavaScript/OOXML is always confirmation-first, bounded, scope-frozen, verified, and
  rollback-aware.
- Terminal answers are receipt-derived and never claim success without the required proof.
- Component or host rollback is immediate through policy and requires no document migration.
