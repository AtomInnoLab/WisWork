# Rapid iteration without a release install

Use the isolated development branch/worktree for code changes. Production Taskpane, Relay and WisWork upgrades remain manual operations; none of the commands below deploy them.

## Local macOS dogfood

```sh
npm ci
npm run dogfood:mac -- --dry-run
npm run dogfood:mac -- --no-launch
npm run dogfood:mac
```

The first build needs Node 22.12+, Rust 1.88+, Xcode command-line tools and network access for locked dependencies/native assets. If Rust 1.88.0 is already installed but the default is older, use `RUSTUP_TOOLCHAIN=1.88.0 npm run dogfood:mac` for this command only; no global default change is needed. Warm builds reuse matching source/output fingerprints; changes to shared packages conservatively rebuild the required editors. A cold build is not expected to meet the 1–3 minute warm-build target. Measure the printed duration on your machine before promising a latency target.

The app is **WisWork Dogfood**, in `apps/shell/release-dogfood`, not `/Applications/WisWork.app`. Its profile is separate from both production and `WisWork Dev`. No production data migration, update feed or file association takeover is allowed. Close the previous Dogfood instance before relaunching to avoid its single-instance lock; never kill the production app to clear that lock. Use test documents. Existing production credentials are not copied. Verified components already installed in that profile can be reused by the existing component manager; optional binaries are not bundled into the app.

**Current authentication limitation:** the existing OAuth callback is hard-coded to `wiswork://`, which belongs to production. Test builds deliberately do not register or consume that callback. A fresh isolated profile therefore cannot complete ordinary browser sign-in yet. Local editor/package tests work independently, but real-model dogfood acceptance requires a dedicated test callback flow and matching authorization-server configuration. Do not work around this by taking over the production protocol or copying production credentials. This first delivery is not a complete authenticated dogfood loop.

Build output includes version, commit, timestamp and mode plus log/artifact paths. Local build logs are not privacy-scrubbed diagnostic exports: inspect them before sharing. Dirty working-tree builds may contain changes beyond the named commit.

## Protocol diagnostics and offline replay

Export the Enhanced diagnostics report, then run:

```sh
npm run diagnostics:replay -- /absolute/path/report.json
```

This is offline protocol validation, not another model request. Keep the exported report associated with the task that failed; an old report cannot diagnose a newer build's run.

Only the latest four captures from the current process are retained; restart clears them. Use `--index 0` through `--index 3` to select an available capture. Capture metadata is separate from the replay result. Recordings are not yet correlated to document/task authority, so use timestamps and original outcome as investigative evidence—not as proof that a particular task caused the capture. `structural-only` is always shown in replay output.

The replay path records bounded, allowlisted protocol structure. It must never export raw SSE bodies, document/user text, JWTs, tool argument bodies, reasoning content or filesystem paths. Unknown or invalid report fields must not be echoed in CLI errors.

Structural replay can test event order, block lifecycle, stop reasons and protocol outcomes. It cannot reconstruct redacted content, simulate every network failure, or prove that a document mutation completed. For document workflows, pair the report with independently verified page/cell/shape results in a disposable document.

## PR Preview

Every PR to `main` gets an arm64 macOS **internal unsigned** build via `Desktop Preview (unsigned internal)`. Download its ZIP from the workflow summary. The ZIP preserves application symlinks and executable permissions. Artifacts are retained for seven days; downloaded apps do not self-delete or automatically expire.

Each PR has its own application name, bundle ID and profile (`WisWork Preview PR123`). The version includes PR number and commit. It has no production update feed or file/protocol registration. Never rely on it as a signed, notarized distribution or an automatic release candidate. It can contain arbitrary PR code: only run a revision you have reviewed, with disposable documents and a test account.

Automatic PR workflows have a read-only token and no signing credentials. A **signed/notarized shareable Preview is a separate follow-up**: approve an immutable, reviewed commit in a protected environment, then sign and verify its artifact using release-quality gates. Do not add signing secrets to `pull_request`, use `pull_request_target` to execute PR code, or publish Preview into the production feed.

## Verification tiers

- While editing: focused unit tests and typecheck for the changed workspace.
- Before submitting: rapid-iteration tests, related workspace suites, privacy/protocol fixtures, lint and changed-file formatting.
- PR: existing full CI plus isolated Preview packaging. Preview packaging success alone does not prove real-model or Office acceptance.
- Release: existing signed/notarized multi-platform pipeline and real install/update/rollback acceptance remain required.

Real-model acceptance should cover: five-slide generation; search then presentation; image insertion; encrypted reasoning and token truncation; incomplete tool arguments; cancellation/retry; missing/corrupt component; Standard/Enhanced switch and restart; failure export/replay. Record screenshots, document counts, tool receipts and a safe report—not only process exit codes. Use a dedicated test account; never commit its credentials.

Rollback: keep production installed and stop the isolated test app. Generated build artifacts may be removed separately after inspecting their exact paths. Preserve the isolated profile if you want settings on the next build; production data does not need restoring.
