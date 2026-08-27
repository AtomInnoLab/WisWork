# Remove Codex and Enhanced mode

## Goal

Remove the optional Codex/Enhanced mode product path from WisWork so the shipped desktop app uses
only the existing Standard agent path and no build, runtime, or UI can download or launch Codex.

## Non-goals

- Do not change the shared `@wiswork/agent-core` or `@wiswork/agent-harness` behavior.
- Do not change Office Taskpane, Relay, model/provider selection, auth, or production endpoints.
- Do not remove generic LaTeX proposal, snapshot, rollback, or close-safety improvements that are
  useful to the Standard path.
- Do not rewrite historical branch names in unrelated archived planning documents.

## Architecture

The Codex feature is a closed optional chain from the account-menu Enhanced mode control through
Shell preload/IPC, a downloaded component manager, a local Codex app-server bridge, and a LaTeX
renderer tool adapter. Remove that entire chain and its release inputs while restoring Standard
mode as the only runtime. Keep generic proposal/snapshot transaction hardening and generic tab
close safety, deleting only their Codex-specific consumers and API surface.

## Global constraints

- No executable download, process spawn, local response server, MCP bridge, or renderer authority
  introduced for Codex may remain reachable.
- Existing Standard AI, LaTeX review/apply/undo, Office, Docs, Sheets, Slides, and Markdown behavior
  must remain unchanged.
- Old `app-settings.json` values such as `agentRuntime: "enhanced"` are ignored; no destructive user
  settings migration is required.
- The base macOS arm64 DMG/ZIP workflow remains intact. Only the optional Enhanced mode job and
  filters/checks are removed.
- Legal notices and license gates must stop reading or publishing removed Codex assets.

## Affected components

- `apps/shell`: remove Enhanced mode account UI, preload APIs, IPC, component manager integration,
  Codex runtime/process/tool bridge, and Codex-specific close coordination.
- `apps/latex`: remove the Codex renderer tool session and Codex-only IPC/runtime branches; retain
  Standard AgentLoop and generic proposal review/apply/undo.
- `packages/codex-bridge` and `packages/agent-runtime`: remove both workspaces because they have no
  non-Codex consumers.
- Root/package manifests: remove workspaces from test/typecheck chains and consumer dependencies;
  regenerate the lockfile mechanically.
- `tools`, legal notices, release docs, and `.github/workflows/package-macos.yml`: remove Codex
  download, eval, schema, signing, license, and Enhanced mode publishing paths.

## Failure handling and migration

This is a fail-closed removal. There is no fallback launch path and no hidden environment-variable
override. An existing downloaded component may remain in a user's private cache as inert data, but
the application has no API or code path to discover or execute it. A later cleanup release may
remove that cache only with an explicit, separately reviewed migration.

## Rollback and release

Rollback is the normal Git revert of this removal commit, but Enhanced mode must not be re-enabled
until its pinned macOS artifact passes the original integrity, signature, Gatekeeper, protocol, and
license gates. Release requires a new WisWork PC build only; Taskpane, Relay, and Manifest do not
change.

## Verification

- A removal-policy test proves no active product/runtime/CI/package/legal source references Codex,
  Enhanced mode, `@wiswork/codex-bridge`, or `@wiswork/agent-runtime`.
- Shell and LaTeX focused/full tests and typechecks pass, including Standard AI, proposal review,
  undo, dirty-close, and tab lifecycle coverage.
- Root format, lint, typecheck, licenses, notices, build, and applicable tests pass.
- The macOS workflow contains only the base arm64 packaging job and no optional component download.
- Final package inspection confirms no Codex/Enhanced mode executable or resource is bundled.
