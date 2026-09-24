# Post-write visual evidence for native existing-deck changes

## Goal and boundaries

Follow the approved PPT Agent plan §6.4/O4/O5: after each confirmed existing-deck native write has passed target readback, capture affected page screenshots and return them to the Agent for visual review. A screenshot is evidence for a page at capture time, not a QA pass. The user has deferred real PowerPoint acceptance (0/20), and this task will not change that claim. Preserve the original plan and the isolated implementation branch; no merge, push or deployment.

Do not make a successful host write appear failed solely because screenshot capture or transport failed. Do not infer a page by index after native replacement changes its host ID. Limit captures to exact IDs and bounded PNG bytes. Do not store screenshot pixels in Office settings.

## Design

Extend structured proposal requests with an optional post-verification evidence callback. The controller runs it only after `execute` and `verify` succeed, isolates errors from the write outcome, and passes a bounded result through the confirmed proposal decision. The final tool result displays captured images to the user through a UI side channel and gives the Agent page IDs, digests and capture status; the Agent uses one-page capture tools to inspect images in model context without exceeding request limits. Each existing-deck skill supplies exact current host slide IDs and rechecks its saved terminal state around the capture. Existing manual review tools remain available and still require an explicit pass/fail judgment.

## Deliverables

### 1. Confirmation evidence channel

- Files: `proposal-controller.ts`, `use-office-agent.ts`, their tests.
- Interface: optional `postWrite` callback and a bounded `ProposalPostWriteEvidence` with page IDs and PNG image content. A confirmed decision may carry captured evidence or a clear unavailable status. `finalProposalExecution` forwards images to UI display, bounded page metadata to the model, and states `qaPassed:false`.
- RED/GREEN: no callback for rejected or failed writes; capture failure leaves the write confirmed and reports unavailable; oversized/malformed evidence cannot enter the Agent result. Commit scoped.

### 2. Existing text and batch captures

- Files: `presentation-existing-editing.ts`, `presentation-existing-batch-editing.ts`, focused tests.
- After successful target readback, capture one page for text/geometry and each affected page for batch. Compare the exact native slide IDs, terminal saved record and target values before and after screenshot. Reuse the existing screenshot validator. Return images without recording visual pass; manual historical review remains separate. Commit scoped.

### 3. Existing image and whole-page captures

- Files: `presentation-existing-image-editing.ts`, `presentation-existing-page-editing.ts`, runtime adapter wiring and tests.
- After successful replacement/undo, capture the actual current native page ID. Image edits retain the page ID; whole-page edits use the inserted/restored ID from the durable record, and staged/discarded phases identify the appropriate page(s) explicitly. Verify saved state and host content before/after capture. Commit scoped.

## Verification and release

Run focused tests, full Office Add-in suite, TypeScript check, ESLint, production build and `git diff --check`; independently review final code for false QA claims, incorrect page selection and misleading success/failure. Record phase progress honestly. Existing records remain readable; no settings migration. The post-write image evidence is session output, not a durable QA verdict.
