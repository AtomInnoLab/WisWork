# Office production recovery release

User request: increment PC one patch, open a PR, deploy Taskpane and Relay. Release the reviewed recovery branch; do not auto-merge, create a production desktop tag, or claim a Mac installer has been published.

## Versioned deliverable

Update `apps/shell/package.json` and its entry in `package-lock.json` from 0.6.69 to 0.6.70. Update `apps/office-addin/public/manifest.xml` and the manifest/build-output version assertions from 0.3.39.0 to 0.3.40.0, including its cache-busting query. No runtime behavior or dependencies change. Verify a clean diff, complete repository tests/typecheck/lint/format, desktop build, and release policy tests. Open the branch PR against verified `origin/main` (867cc93); retain the worktree for PR feedback.

## Deployable deliverables

1. Build the locked Relay release, run its full test suite/clippy, and verify production paths and service state. Stop Relay for a consistent protected binary/SQLite backup, accounting for any journal/WAL sidecars. Validate the copy with SQLite quick_check and schema version. Run the new binary against a staging database copy; only then atomically replace `/opt/wiswork-relay/wiswork-relay` and restart `wiswork-relay`. Verify local/public health and binary checksum. No database migration, config/auth change, or source payload logging.
2. After tests finish (tests also write `dist`), build Taskpane for `https://office.8-216-134-194.sslip.io` with the release commit build ID and existing persistent-pairing flag. Back up `/var/www/wiswork-office-addin` entry/manifest, install assets before atomically switching entry/manifest, retain older assets and `components`. Verify public manifest, entry, bundle checksum, build ID, and Relay health. No Nginx changes needed.

Independent release review checks version files, capabilities, deployment and rollback. Capture artifact checksums and backup locations in the PR after successful cutover. Old PCs continue negotiated old capabilities. PC recovery requires PC 0.6.70; only the new DESIGN.md editing/synchronization capability additionally requires a new approved pairing. An active Relay restart interrupts current sessions; users reconnect afterward.

## Rollback and acceptance boundaries

If a health or artifact check fails, restore the recorded binary/static entry and restart the previous service. Before downgrading capability-bearing clients, disable pairing resume in reverse order (Taskpane `VITE_WISWORK_OFFICE_PAIRING_RESUME=0`, PC `WISWORK_OFFICE_PAIRING_RESUME=0`, Relay `WISWORK_RELAY_PAIRING_RESUME=0`), preserving IndexedDB, the encrypted PC binding store, and the production database. This release has no schema change, but old clients do not understand new DESIGN capability records. Never restore an old database over newer bindings/revocations just to roll back a binary; verify compatibility before re-enabling resume. Stop+backup and staging checks are done with protected permissions; never print DB content, tokens, or request payloads. Keep backup material outside public web roots.

Automated passing checks do not constitute real Mac PowerPoint full-deck acceptance. PR CI/Preview status and production service deployment status must be reported separately. PC release publication follows merge and the existing Desktop Release workflow, not this PR creation.
