# Office PowerPoint and WisWork Slides parity implementation plan

## Goal and architecture

Make Office PowerPoint Taskpane consume the same display-safe presentation-agent vocabulary and visual primitives as desktop Slides while preserving separate host adapters and document authority. Shared code will live in `packages/ui`; presentation runtime truth remains in `agent-core`, Slides, and Office.

Global constraints: no raw renderer colors, no document authority in UI packages, no silent Standard/Enhanced fallback, no removal of proposal/transaction/readback/rollback controls, and no false parity for unsupported Office.js features.

## Task 1 — Shared presentation-agent view primitives

- Modify `packages/ui/src/index.ts` and add focused components/types under `packages/ui/src/`.
- Consume display-safe messages, grouped activities, progress notices, errors, and recovery callbacks.
- Add component tests proving exact state rendering, accessibility, bounded details, and narrow-layout class contracts.
- RED: Office and Slides parity fixtures cannot render through a common contract.
- GREEN: both host fixtures render identical semantic states from the same model.
- Commit as an independently reviewable UI foundation.

## Task 2 — Office PowerPoint display adapter and responsive UI

- Modify `apps/office-addin/src/App.tsx`, `styles.css`, and the Office timeline/view-model modules.
- Map Office session events to the shared contract; replace stacked status chrome with grouped steps and desktop-consistent messages, empty state, composer, confirmation, error, retry, and stop behavior.
- Localize new visible strings through `@wiswork/i18n`; avoid hard-coded English in the PowerPoint experience.
- Add jsdom/static markup and CSS policy tests at wide, narrow, dark, forced-color, and reduced-motion states.
- RED: current Office markup fails the desktop semantic parity fixture.
- GREEN: Office PowerPoint passes the shared fixture and existing Word/Excel tests remain unchanged.
- Commit separately.

## Task 3 — PowerPoint creation workflow parity

- Modify `apps/office-addin/src/skills/powerpoint/powerpoint-skill.ts`, `agent/host-runtime.ts`, and `agent/use-office-agent.ts` only where needed.
- Add a PowerPoint-specific creation prompt/contract matching desktop sequencing: context, material clarification, plan, research, bounded construction, verification, receipt.
- Reuse existing Office proposal and presentation-verification authorities; unsupported desktop operations must be represented as capabilities or safe alternatives, not invented tool calls.
- Add Standard and Enhanced integration fixtures for new-deck and existing-deck workflows, including cancel and stale proposal paths.
- RED: end-to-end Office PowerPoint fixture lacks plan/build/verify parity or terminal truth.
- GREEN: both runtimes reach the same semantic receipt through Office adapters with explicit confirmation.
- Commit separately.

## Task 4 — Compatibility and release verification

- Run complete Office, Slides, agent-core, agent-harness, agent-runtime, presentation-verification, UI, and i18n tests.
- Run root typecheck, changed-file lint, Prettier, `check:theme-colors`, Office/Slides/Shell production builds, and presentation E2E/artifact restoration.
- Inspect the final diff for secrets, raw colors, protocol leakage, document-authority movement, and unrelated changes.
- Record any environment-only E2E limitation without claiming it passed.
- Obtain independent code review before merging.
