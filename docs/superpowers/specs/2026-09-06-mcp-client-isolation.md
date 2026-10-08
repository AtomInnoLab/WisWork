# MCP client isolation

Approved direction: isolate MCP connection initialization state and RPC request IDs per runtime client without weakening document authorization or replay protection.

The shared gateway remains the document authority. Transport sessions must not grant document access: existing turn capabilities, document identity, confirmation requirements and call IDs remain authoritative. Multiple native threads must be able to initialize against the shared endpoint independently. A repeated request inside one transport session must still be rejected. Cross-client initialization must not reset another client's state or pending work.

Scope: trusted gateway transport, protocol client integration where necessary, bounded diagnostics, and regression tests. No production deployment, authentication reset, global cache clearing or silent permission expansion.

Acceptance: reproduce the second-client failure first; demonstrate independent initialization and interleaved requests; reject invalid session IDs and replay; retain untrusted document-server behavior; verify cleanup and resource bounds. Run the real pinned runtime against a local fake upstream where practical, without user credentials. Full repository checks and independent review precede delivery. Real PowerPoint acceptance remains a separate explicit verification, not inferred from unit tests.

Rollback: revert the scoped transport fix and its integration together. No persistent data migration.
