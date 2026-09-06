# MCP client isolation implementation

Base: origin/main 2692b495. Specification: ../specs/2026-09-06-mcp-client-isolation.md.

1. Reproduce and repair trusted multi-client transport in packages/codex-bridge/src/mcp-server.ts and tests/mcp-server.test.ts or tests/dynamic-mcp-gateway.test.ts. Choose a bounded session mechanism compatible with the pinned client; retain all document capability checks. Observe RED before implementation and GREEN after. Include initialization, interleaving, same-client replay, invalid session, cleanup and bounds. Commit this independently testable unit.
2. Validate integration with the pinned runtime and Shell. If the protocol requires client changes, keep them in the same scoped unit and test them. Preserve safe diagnostic codes without exposing prompts, credentials or document content. Independently review the complete diff; fix substantive findings.
3. Run formatting, lint, whole-repository typecheck, tests and licenses. Build Shell after any main-process changes. Report exact evidence and limitations. Do not deploy or replace installed applications without further instruction.

No durable migration. Roll back the scoped unit together. Do not solve client collisions by globally clearing IDs or accepting repeated initialization in an existing session.
