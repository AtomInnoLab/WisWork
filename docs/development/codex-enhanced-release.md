# Codex Enhanced release and incident runbook

Codex Enhanced ships as one seven-host compatibility set for LaTeX, Slides, Docs, Sheets, Office Word, Office Excel, and Office PowerPoint. Standard remains the default and recovery path. Release only after native macOS arm64, macOS x64, and Windows x64 jobs pass from a clean `npm ci`, including a pinned 0.147.0 same-package protocol read and post-package proof that the base application contains no Codex executable.

## Rollout checklist

1. Confirm the reviewed component manifest, hashes, publisher identity, notices, and mirror/fallback byte identity.
2. Run `npm run test:codex-rollout`, `npm run test:enhanced-seven-host`, the Slides render E2E, full tests/typecheck/build, license/notices checks, and package inspection.
3. Exercise download, restart-only activation, launch, update beside the active version, removal, offline failure, corrupt bytes, fallback, signature rejection, runtime crash, global switch, every per-host switch, and the raw Office switch on native CI.
4. Enable all seven host switches only after the entire gate is green. Do not partially advertise availability.

## Incident rollback

Rollback in this order: disable the raw Office switch; disable the affected per-host switch; disable global Enhanced; remove the component-version allowlist entry; then revert code. Every switch fails closed and restores Standard behavior for new sessions without replaying an Enhanced turn. Cancel undispatched work and preserve receipt truth and history for applied work.

`applied_unverified` means a mutation may have applied but proof is unavailable: do not retry automatically; retain rollback and request review. `write_pending_quarantined` means the write may still be running: freeze further edits until reconciliation or reload. Never rewrite either state as verified. Telemetry contains only closed host/phase/outcome enums and is non-authoritative; collector failure must not affect a task.

## Removal and recovery

Removal targets only the resolved app-private immutable component directory and activates after restart when required. Revalidate path, hash, version, and publisher before every launch. A failed update leaves the old verified version active; an invalid or unavailable component starts no Enhanced task and never falls back mid-turn.
