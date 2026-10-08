# First-delivery verification

Implementation base: `195867d0` (origin/main). Implementation head: `9e4c1b6a`.

## Passed locally

- `TMPDIR=/private/tmp RUSTUP_TOOLCHAIN=1.88.0 CARGO_NET_OFFLINE=true npm test`: exit 0; 6,801 Vitest tests passed and 53 Rust tests passed, plus root Node policy suites. Optional integration tests retain their existing skips.
- `npm run test:rapid-iteration`: 14/14 passed.
- `npm run lint`: exit 0, 13 existing warnings and no errors.
- `npm run typecheck`: exit 0 across all workspaces.
- `FORMAT_BASE_REF=origin/main npm run format:check`: exit 0.
- `git diff --check`: exit 0.
- `npm run diagnostics:replay -- packages/codex-bridge/tests/fixtures/protocol-redacted-max-tokens.json`: production parser emits `response.incomplete`, explicitly `structural-only`.
- `npm run dogfood:mac -- --dry-run`: prints isolated identity, provenance, selected builds and paths without packaging or launch.

Independent unit and final integration review completed. Important findings fixed: stale/mid-build cache publication, omitted component-manifest input, native source override mismatch, arbitrary numeric diagnostic content, and supported metadata distortion. Original stream outcome is distinct from simulated replay outcome. Task association is explicitly unattributed and deferred.

## Environment findings

The default Rust is 1.85.0, below current dependency requirements. Rust 1.88.0 was already installed; the final test command selects it only for that process. No global toolchain change was made. Required host crates were fetched with bounded network timeout. macOS temporary directory aliases required canonical `TMPDIR=/private/tmp` for the existing tab-session assertions. Loopback tests needed execution permission outside the restricted sandbox.

## Not verified / not shipped

- Actual Dogfood `.app` packaging or launch: attempted, but initial dependency/download work stalled and was stopped. Electron 43.4.0 package is installed, but its `dist/Electron.app` binary is still missing. No usable app artifact is claimed.
- Warm-build 1–3 minute target: not measured.
- Preview workflow: configuration and isolation tests passed locally; not executed on GitHub. Only unsigned internal artifacts are implemented. Signed/notarized sharing requires a protected follow-up.
- Authenticated model loop: fresh test profiles cannot use the production-owned `wiswork://` callback. Dedicated callback/server configuration is required. Production credentials were not copied.
- Real UI, Office and live-model acceptance: not executed in this delivery.
- No production app replacement, data migration, service deployment, remote push or PR creation. Original checkout remains unchanged. Branch/worktree retained for continuation.

Detailed command output for this local run: `/private/tmp/wiswork-rapid-tests-rust188.log`. This temporary log is not a portable CI artifact or a scrubbed diagnostic export.
