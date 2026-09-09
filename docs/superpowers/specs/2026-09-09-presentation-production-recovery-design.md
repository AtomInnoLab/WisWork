# Presentation Production Recovery

## Goal

Make the PC Slides agent reliably complete a DESIGN.md-driven deck: research and validate assets before production, build contract-bound pages, screenshot-review and repair each batch, and always release host transactions.

## Boundaries

- Applies to WisWork PC Slides and the shared image-search backend.
- Keeps the schema-version-1 `PresentationDesignContract` and legacy-plan reader compatibility.
- Does not add a second planning format or relax PowerPoint/Slides write authorization.
- Does not expose API keys, upstream response bodies, private chain-of-thought, or unsafe image URLs.

## Design

1. **Search configuration and errors.** Persist the optional SerpApi key with Electron `safeStorage`; environment configuration remains a compatible override. Search returns a typed failure when configuration, authorization, quota, timeout, response parsing, or all providers fail. Only a confirmed successful empty result is rendered as “0 images”. A settings/diagnostic surface can save, clear, and test the key without reading it back.
2. **DESIGN.md readiness.** The UI does not publish an empty DESIGN.md timeline item. `plan_deck` accepts only a complete ready contract. Required assets must be `ready` or have a declared ready fallback before prototype production. The complete rendered contract is persisted and becomes the only source used by subsequent build and review calls.
3. **Stable production binding.** `build_deck` consumes page indexes and verifies normalized contract fields instead of brittle byte-for-byte title punctuation. It rejects real semantic drift but accepts typography-equivalent punctuation and whitespace. Production tools receive the contract revision and slide acceptance IDs in their output.
4. **Repairable review loop.** Screenshot capture is separate from visual judgment. A failed visual review keeps only that page pending and permits the bounded low-level repair tools needed to fix it. A successful re-capture/review clears that page. Screenshot unavailability reports a retryable capability error instead of silently hanging.
5. **Safe editing and transaction cleanup.** Prefer focused element tools for text repair; raw scripts retain syntax validation and actionable errors. Every mutation claim and pending call is settled in `finally`/timeout/cancel paths so one failure cannot leave `tool_call_in_progress` blocking later pages.

## Success Criteria

- Missing/invalid/exhausted SerpApi configuration produces an explicit user-visible failure category, not a successful zero count.
- Broad image queries return provider results when a configured provider succeeds.
- No DESIGN.md card appears until the complete contract exists; opening it shows brief, narrative, visual system, pages, assets, acceptance rules, revision, and status.
- Production cannot begin before contract and asset readiness; after readiness, equivalent titles do not fail binding.
- Prototype pages can be screenshot-reviewed, repaired, re-reviewed, and followed by remaining batches.
- A failed, timed-out, or cancelled mutation cannot cause a subsequent read or write to return `tool_call_in_progress` after settlement.

## Compatibility, Security, and Rollback

- Old environment-based `SERPAPI_API_KEY` remains supported and takes precedence for automation.
- Older saved decks without a structured snapshot still use legacy `plan_deck` parsing; they must re-plan before mutation.
- Secrets are encrypted at rest when `safeStorage` is available and never returned to the renderer.
- Rollback is the single feature commit; persisted secret state can be cleared from settings and does not alter deck files.
