# Tool lifecycle delivery and cancelled-turn isolation

Approved direction: the user approved continuing after the audit and proposal to unify semantic lifecycle delivery and isolate by turn/call identity.

## Goal and boundaries

Every authenticated, known semantic tool attempt has one visible start and one terminal outcome, including router rejection, declined proposals, execution errors, and cancellation. Office enhanced sessions must support more than 32 retrieval attempts without confusing a display budget with a model-turn budget. Cancelled desktop reads must not publish results into a subsequent run.

Do not execute observation frames, expose carrier credentials/raw error bodies, weaken authorization/readiness checks, automatically retry writes, change model carrier exec/wait presentation, release a desktop version, or deploy Taskpane.

## Design

Use the existing dynamic gateway tool-start/tool-complete events as the authoritative lifecycle source. Add opaque per-grant turn identity and bounded stable error metadata, never raw tool input or output, and preserve existing consumers through additive metadata. Existing host executors can enrich a matching call's display, but must not create duplicate starts or reverse a terminal outcome.

Office projects lifecycle events onto the current request/session generation, and scopes deduplication to turn and call. It retains strict byte, event, and identity bounds while aligning the lifecycle call budget with the document session's 1,024-call bound. Display enrichment stays bounded and cannot own execution. Existing retrieval frames retain their type and shape; other semantic observations use a new event type that older stream parsers ignore. This preserves mixed-version operation without injecting unsupported metadata into upstream model requests.

Desktop native registration forwards gateway lifecycle, propagates read abort signals, clears pending work before terminal publication, and rejects stale receipts. Generic renderer callbacks and Slides' local harness must deduplicate execution events without dropping router/proposal outcomes for tools that never execute locally. Turn/call bindings and cancellation tombstones must be bounded.

## Acceptance

- A real proxy + document session + Office transport run can perform 33+ retrievals and finish with exactly one start/end per call.
- A read rejected by a pending write displays an error without executing the read.
- Normal Office reads/writes/searches retain their output/preview with no duplicate cards.
- Cancelling an old PC read, starting another run, and resolving the old read cannot dispatch a successful receipt to the new run.
- Declined/expired Slides proposals are observable without executing the proposed mutation.
- Replay, foreign/stale generation, oversized metadata, and untrusted tool names remain rejected; model/carrier internals stay private.

## Verification and rollback

Start with regression tests for each audited failure, then run adjacent suites, full repository tests/typecheck/lint/format/licenses, and affected builds. Independently review the complete changes including privacy and lifecycle ordering. No live Office/Mac or production claims from mocked tests.

Changes remain on the existing local fix branch. There is no data migration; rollback is a scoped revert of this lifecycle follow-up. Keep the preceding asset/readiness fix intact. Publishing and deployment require a separate user decision.
