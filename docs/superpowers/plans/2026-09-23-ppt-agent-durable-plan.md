# PPT durable planning / batch 3

Base 34a04f3. Continue approved v1.3 O1 project model in existing isolated worktree. High assurance workflow and ponytail remain applicable; no new dependencies, deployment or merge.

## Architecture

A browser-safe PresentationPlan schema stores the brief, source/claim ledger, shared style and ordered slide tasks. A document-bound CAS plan record persists independently of compilation. Compilation of planned projects must match the saved revision and deterministic content/style contract; each compile receipt snapshots the plan so later planning changes cannot rewrite its evidence. Legacy unplanned projects keep working.

## Interfaces

`@wiswork/pptx-engine/presentation-plan`: PresentationPlan, PRESENTATION_PLAN_SCHEMA, parsePresentationPlan, presentationPlanClaims(plan), assertDeckMatchesPresentationPlan(deck,plan).
Plan: {version:1,projectId,title,brief:{objective,audience,language,minutes,requiredContent:string[],constraints:string[]},sources:[{id,title,uri,locator?,excerpt,asOf?}],claims:[{id,statement,type:'fact'|'quote'|'calculation'|'judgment'|'assumption',sourceIds:string[],confidence:'high'|'medium'|'low',reviewStatus:'needs_review',asOf?,jurisdiction?,calculation?:{formula,inputs:string[],unit?,currency?}}],style:PresentationStyle,slides:[{id,title,purpose,claimIds:string[],layout:'cover'|'content'|'comparison'|'process'|'chart'|'summary',requiredAssets:string[],acceptanceCriteria:string[]}]}. <=192 KiB UTF-8. ID max80; slides 1..32; sources/claims <=256; source uri<=500, locator<=200, claim sourceIds<=3. Source URI is metadata, never fetched. Fact/quote/calculation require sources; calculations require formula+inputs metadata (no eval). All cross refs/IDs unique and valid. Unverified is the only machine-set review status. Claim mapping: text=statement, source=source URIs joined(' ; '), locator=nonempty source locators joined(' ; '), absent if none. For zero-source judgment/assumption map source='未核验：'+type (to preserve honest footer). Compile contract validates project ID/title/style, exact ordered slide IDs/titles/claimIds and exact mapped deck claims; no fake content/visual QA.

Store API:

- plan(projectId,documentId): PresentationPlanRecord|undefined
- savePlan(projectId,documentId,expectedRevision,plan): record {version:1,projectId,documentId,revision,plan,inputDigest}
- CAS revision0 creates; same payload retry for expected=current-1 returns same record; stale changed payload => revision_conflict. Current revision same unchanged payload returns same record.
- begin(...,deck,planBinding?:{revision:number,plan:unknown}) retains optional binding in receipt with digest verification. Legacy receipts stay valid. Existing request conflict includes binding changes; complete preserves it. Snapshot must match saved plan in service.

Service additions:

- save_plan {documentId,projectId,expectedRevision,plan} => {projectId,revision,plan}
- get_plan {documentId,projectId} => same or not_found
- compile optional planRevision; required and equal current plan revision for new request in a planned project; duplicate request uses its saved binding and original input. resume uses saved binding. errors invalid_plan/plan_mismatch/revision_conflict.
- status: existing response gets optional requestPlanRevision:number for latest receipt binding and optional plan:{revision,value:PresentationPlan}; plan-only project returns status:'planned', slides from plan, history:[], no latestRequestId/checks. Existing compiled artifacts stay separately visible if plan later changes; UI labels them separately.

## Reviewable units

1. Schema/contract + fixture/tests: new pptx-engine presentation-plan.ts, test and export; agent owns. TDD invalid refs, forged review status, invalid calculations, exact deck mapping, size bounds, legacy deck unaffected. Commit.
2. Atomic plan store + receipt snapshot/tests: project-store presentation-store.ts only; agent owns. TDD CAS, identical retry, changed stale input, foreign documents, symlinks/corruption, plan snapshot retained with old receipts. Commit.
3. Root service + Agent tools + project card/controller: add save/read planning tools, recover plan as local JSON and model output, lifecycle/document checks; compile checks saved plan revision; status supports precompile plan and avoids implying latest artifact reflects a newer plan. Tests real save/reload/compile/revise/retry flow, user-readable revision conflicts. Commit.

## Verification and release

Independent unit reviews then broad integration review. Fresh full npm test and typecheck, changed TS lint, production builds (serialize builds and tests due shared build fixtures), diff check. Add stage report with achievements and next steps. Plan file is additive; older code ignores it and optional receipt fields, retains artifacts. Rollback to older PC does not enforce new plan gate, so keep paired PC/Taskpane versions aligned. No UI permission dialog for saving plans; native document import confirmation remains separate. No attachment transport or page-level execution claims in this batch.

## Execution evidence

- Schema/fixtures: 6868f4f; shared expanded-text budget fix903e731; mandatory UTF-8 projection bound d0f8f71. Both budget issues were independently reproduced and fixed with RED→GREEN boundary cases. Deck helper extraction preserves legacy behavior.
- Store: ba337b8; downstream strict optional-property compatibility fix cf3e2a5. CAS, immutable binding and legacy receipt tests pass.
- Service RED: 3 planning cases rejected as unsupported; GREEN: save/read, planned status, revision/contract enforcement and old-request snapshots pass.
- Runtime RED: planning tool absent and plan_revision rejected; GREEN: routed lifecycle-safe tools and optional compile revision.
- Project UI RED: plan-only project rejected/incorrectly labeled; GREEN: plan-only view and separate current-plan/old-artifact display.
- End-to-end: actual8-page PptxGenJS output recovered and parsed after persisted-plan reload; advancing the plan does not mutate old receipts or compiled bytes.
- Independent broad review found only the two budget issues above; changed scopes re-reviewed clean. Final verification recorded in the stage report.

- Build-test diagnosis: full-suite setup hit the10s default hook timeout; isolated4 build assertions passed in25.34s total (three production builds). The fixture now has a local30s beforeAll budget; all578 addin tests pass without changing assertions or global timeouts.
