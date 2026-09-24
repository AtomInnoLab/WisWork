# Chart backup quota reclamation

Goal: free PC backup quota only for chart changes whose journal is terminal (`cancelled` or `undone`). Do not remove backups for an applied, pending, or uncertain change. Release is an explicit confirmed tool action; no automatic retention timer.

Design: the PC backup service adds an exact-scoped release request under its per-document lock, with durable tombstone for idempotent retry. The Office journal records `backupReleasedAt` only after the PC acknowledges release. The saved-changes UI offers release only for terminal chart records without a release receipt. A failed request leaves the journal and quota unchanged or safely retryable.

Files: PC backup service, presentation service routing and shell tests; chart record/history/UI and Office tests; `powerpoint-skill.ts` release proposal and host-runtime exposure. No migration of existing backups.

1. PC release service: test ready, exact metadata, idempotent retry, active upload rejection and quota reuse; implement scoped deletion plus durable receipt. Scoped commit.
2. Chart record and UI: test terminal-only release receipt, CAS persistence and available action; implement validator/history action. Scoped commit.
3. Office release tool: test confirmation, PC failure, host/document drift and receipt finalization; execute only for terminal chart records. Scoped commit.
4. Independent review; full Office and focused shell verification, typecheck, lint, build. Document real-host acceptance status and quota limitations.
