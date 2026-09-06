# Rapid iteration — first delivery

Base: origin/main 195867d0. User approved implementation and isolated worktree.

## Contract

Implement local macOS dogfood builds, safe protocol replay, and isolated PR previews. Never replace the installed WisWork, migrate production user data, deploy services, or enable production auto-update for test builds. The 1–3 minute warm-build target is a measurement goal, not an acceptance claim without measurement. Signed distributable previews require a trusted approval boundary; untrusted PR builds receive no signing credentials.

Replay captures only enumerated protocol structure and bounded numeric metadata. It must not store JWTs, user/document text, filesystem paths, tool argument bodies or reasoning content. Replay diagnoses protocol-state failures; it cannot reconstruct removed content or faithfully reproduce arbitrary network failures.

## Units

1. Dogfood: tools/dogfood-mac.mjs and tests; isolated packaging config and shell startup identity. First build all required modules, subsequent builds conservatively invalidate affected modules and shared dependencies using content fingerprints. Native assets must be present and verified by existing mechanisms. Build only a host-architecture .app, no DMG/ZIP, no production mutations; dry-run and no-launch support. Reuse only independently verified Enhanced cache; do not bundle optional components. Print commit/version/time/mode and artifact/log locations.
2. Protocol replay: codex bridge recorder/replayer, fixtures and tests, CLI. Connect safe recording to the existing real translation path and diagnostics export where possible; use strict bounded allowlists. Cover encrypted/redacted reasoning and max_tokens, incomplete tool JSON, malformed streams, schema/privacy rejection. Offline CLI uses production protocol logic rather than a hard-coded result lookup.
3. Preview: isolated per-PR identity and build metadata, no release feed/protocol takeover or production migration. Automatic PR workflow with read-only token, expiring artifacts and no secrets; clearly label unsigned CI previews. Document protected, signed shareable preview follow-up if signing infrastructure cannot safely be reused. Add deterministic workflow/config validation tests.
4. Integration: npm entry points, CI focused gates, developer guide with exact commands, limitations and rollback. Fresh targeted tests/typechecks/lint/format and independent full-diff review before completion.

## Verification and rollback

Each implementation unit starts with a failing test for testable behavior and provides verification evidence. Record baseline failures separately. Build/run the local package only when prerequisites allow, without touching production app state. Rollback means remove only generated test artifacts and revert this branch; production installation/data remain untouched. Do not claim UI validation, signed preview delivery, production recording or full-suite success unless actually verified.

## Continuation: real packaging findings

Observed: `npm run notices` exits successfully with zero Rust crates when unfiltered cargo metadata requires uncached foreign-platform dependencies; the same locked metadata succeeds offline with `--filter-platform aarch64-apple-darwin`. Add an explicit target-scoped strict notices option for Dogfood and Preview, preserving existing release behavior until separately reviewed. Files: `tools/gen-third-party-notices.mjs`, a testable metadata helper and Node tests, `tools/dogfood-mac.mjs`, `.github/workflows/desktop-preview.yml`, developer guide. Require nonempty metadata and fail before publishing notices on metadata error. Test missing crates/command failure, valid target selection and source preservation; run actual generation and package verification. Do not silently ship incomplete notices. Authentication changes await a separately approved callback/server configuration design.
