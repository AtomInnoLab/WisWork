# Publish reviewed fixes and deploy Taskpane

## Scope and architecture

User request: push the existing fixes and update the deployed Taskpane. Preserve PC 0.6.70 and Taskpane manifest 0.3.40.0; do not release a PC installer, merge the PR, deploy Relay, or change service configuration. The named isolated branch is `codex/fix-office-production-recovery`; existing PR #202 targets main. Remote head was verified as 73a5476 with four unpublished repair commits through 57f7cbf and no remote divergence.

Taskpane is an existing Vite static build served by Nginx from `/var/www/wiswork-office-addin` at `https://office.8-216-134-194.sslip.io`. Existing Relay and PC versions retain their negotiated capabilities; installing this Taskpane alone must not enable lease renewal on old components. No database, authentication policy or stored pairing data changes are included.

## Deliverables and verification

1. Verify the exact branch to publish: fresh complete `npm test`, `npm run typecheck`, `npm run lint`, `npm run format:check`, and independent Taskpane compatibility review of 73a5476..HEAD. Any source changes require renewed affected validation. This release record is the only new file; no new runtime behavior requires a new RED/GREEN cycle. Commit this record as `docs: record Taskpane publication plan`, then fast-forward push to the existing remote branch without force and verify GitHub's head SHA.
2. After tests finish (build-output tests also write dist), build `apps/office-addin/dist` with the exact origin, Relay transport, persistent-pairing flag 1, and the published commit's 12-character build ID. Require a valid generated manifest, fixed WSS CSP, no source maps, complete local assets and matching manifest version. Do not reuse a build made before tests finish.
3. Back up the exact existing static root into a new protected directory outside the web root. Install all new content-addressed assets before atomically replacing taskpane.html and manifest.xml, retaining old assets and any other existing content. Verify public entry and manifest SHA-256 against local files, fetch and compare every built asset, confirm build ID in the entry bundle, and verify Nginx/Relay remain healthy. Record backup path and artifact hashes in the local protected deployment receipt; report deployed build and existing PR to the user.

## Rollback, security and boundaries

No migration or service restart. Never copy secrets, environment files, private keys or database contents into the web root or logs. Use only the established deployment origin and existing flags. If post-cutover checks fail, atomically restore the backed-up entry and manifest; old assets are retained for those entries. Do not roll back Relay or the database. Deployment proves the published static artifact is reachable, not that a real Mac PowerPoint full-deck run has passed. Active task panes must be reopened to load the new entry. PC-side fixes and negotiated renewable leases require a separately authorized PC/Relay rollout.

## Pre-publication evidence

- Fresh `npm test` exited 0: 7,735 Vitest tests passed, 21 skipped; included native tests passed. Fresh locked Relay suite exited 0: 34 unit tests and 56 integration tests passed. Full workspace typecheck completed successfully.
- Fresh full lint and format check exited 0; lint retained 13 existing warnings and no errors. The release record passed explicit Prettier and whitespace checks.
- Independent compatibility and release-plan review found no blocking issues. Unknown lease capabilities are filtered by old Relay, stored pairing grants are not broadened, and diagnostic errors map to previously supported vocabulary.
- Existing public manifest matched the local deployment checksum; existing bundle identified build 73a5476380ae and persistentPairing:true. Nginx and Relay were active; the configured public `/office-relay/health` returned `ok`. `/taskpane.html` returned `Cache-Control: no-store`.
