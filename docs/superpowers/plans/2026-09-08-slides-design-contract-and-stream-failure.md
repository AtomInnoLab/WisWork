# Slides design contract and stream failure plan

Goal: make DESIGN.md visibly editable inside WisWork PC and stop deterministic Enhanced stream failures from retrying and duplicating output. Keep existing local slide editing and Office compatibility unchanged.

Architecture: the existing `.design.md` sidecar remains the single persisted design document. Slides main/preload expose bounded read/write IPC; AiPanel edits that document and supplies its latest value to the Slides skill and screenshot QC. The Responses bridge reports the prepared turn id with deterministic protocol failures so the shell terminates only the matching turn.

Constraints: no new dependencies; validate IPC payloads; preserve untitled-deck pending sidecars; do not terminate unrelated concurrent document turns; do not queue or display a message unless the Agent run accepted it.

Files:

- `apps/slides/src/main/presentation-design-sidecar.ts`: read and update the current design document.
- `apps/slides/src/main/ai-ipc.ts`, `apps/slides/src/preload/index.ts`, `apps/slides/src/shared/ipc.ts`: bounded design-document IPC.
- `apps/slides/src/renderer/ai/AiPanel.tsx`, `apps/slides/src/renderer/styles.css`: editor UI and live context binding.
- `apps/slides/src/renderer/ai/slides-skill.ts`: include the edited contract in Agent context.
- `packages/codex-bridge/src/types.ts`, `packages/codex-bridge/src/local-server.ts`, `apps/shell/src/main/codex-turn-resolver.ts`, `apps/shell/src/main/codex-engine.ts`: turn-correlated deterministic protocol failure.
- Relevant tests: sidecar round-trip, edited context, concurrent-turn failure routing, and rejected-send behavior.

Acceptance:

1. A saved or untitled deck can open, edit, and save DESIGN.md in PC.
2. The next Agent turn and visual QC use the edited text.
3. A deterministic stream protocol rejection settles the matching turn immediately, even with another document turn active.
4. A busy Agent does not create an undelivered user bubble or clear the draft.
5. Targeted suites, Slides suite, typecheck, lint, format, and builds pass.
