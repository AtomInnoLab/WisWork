# Presentation asset planning without fixed count quotas

The user approved the recommended minimal repair after the count audit: remove the whole-contract 120-asset schema quota, the per-slide 20-asset-reference quota, and the legacy per-page four-image-query quota. This does not expand PC `build_deck` from a single focal image to a multi-image layout API, change attachment/search-result quotas, or authorize release/deployment.

## Design and boundaries

Update the shared model-facing contract schema and parser in `packages/agent-core/src/presentation-design-workflow.ts`, then align the legacy `plan_deck` schemas exposed by PC Slides and Office PowerPoint. Remove arbitrary item-count limits only for asset inventory, asset references and image queries. Keep arrays validated as arrays, every item validated for type/nonempty text/length, and all other plan limits unchanged. Do not substitute a larger fixed quota.

Existing ingress byte limits, source download size/pixel checks, searched-URL authorization, search pagination/rate limits, screenshot and geometry bounds, approval/stale-state protections and visual acceptance remain unchanged. This is a planning-quota relaxation, not a claim that arbitrary-size decks fit into one request or that the native host has unlimited memory.

Old valid contracts stay valid with no schema version or persistence migration. New larger plans remain subject to existing request/document resource budgets. Reverting the scoped commit restores the old quotas; users with larger plans would then need to reduce or split those plans. No production data is rewritten.

## Acceptance and risk ownership

Main agent owns shared parsing/schema and final verification; an independent implementer owns the PC/Taskpane exposed-schema alignment and host-level regression coverage. Independent review checks schema/runtime agreement, preservation of per-item validation and safety bounds, and absence of unrelated layout/protocol changes. Tests must first reject the previously capped scenarios, then accept more than 120 assets, more than 20 references and more than four legacy queries without truncation. Invalid items and unrelated limits must still reject. Full repository tests, typecheck, relevant builds, lint/format and review precede local completion. PC/Taskpane publication remains a separate user decision.
