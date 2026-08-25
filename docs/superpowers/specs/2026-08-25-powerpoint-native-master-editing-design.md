# PowerPoint Native Master Editing Design

## Goal

Restore useful `edit_slide_master` behavior on PowerPoint for Mac without importing modified master packages. The tool will use PowerPointApi 1.10 native objects, create one confirmation for the complete declarative program, verify semantic readback, and recover only transaction-owned changes.

Success criteria:

- A user can change master backgrounds, theme colors, and background inheritance with one request and one confirmation.
- Solid, gradient, pattern, and picture/texture master backgrounds are supported when PowerPointApi 1.10 supports them.
- Mac never calls the OOXML master-package replacement path.
- Writes fail before mutation when the captured master/layout state is stale.
- Every changed property is read back and compared semantically.
- Partial execution restores confirmed transaction-owned properties or returns an explicit recovery/uncertainty error.

## Non-goals

- Claiming that native Office.js covers every possible master OOXML mutation.
- Editing placeholder definitions, font schemes, arbitrary relationships, or unsupported master XML on Mac.
- Silently falling back from a master edit to slide-local edits after confirmation.
- Keeping the old raw-XML contract under the same tool name.

## Capability model

Add a bounded read tool `inspect_slide_masters`. It returns:

- PowerPointApi 1.10 availability;
- master IDs and names;
- layout IDs and names;
- master and layout background fill summaries;
- layout background-following flags;
- the supported theme-color slots and current values.

Results are bounded to 32 masters, 128 layouts, and the fixed Office theme-color slot set. Unsupported or oversized decks return explicit bounded-read errors.

`edit_slide_master` becomes a native declarative program with `version: 2` and 1–32 operations:

```json
{
  "version": 2,
  "operations": [
    {
      "op": "set_master_background",
      "master_id": "master-id",
      "fill": { "type": "solid", "color": "#000000", "transparency": 0 }
    },
    {
      "op": "set_master_theme_color",
      "master_id": "master-id",
      "theme_color": "Light1",
      "color": "#FFFFFF"
    },
    {
      "op": "set_layout_background_following",
      "master_id": "master-id",
      "layout_id": "layout-id",
      "follow_master": true,
      "show_master_graphics": true
    }
  ]
}
```

Operations:

1. `set_master_background`
   - `solid`: HTML RGB color plus transparency 0–1.
   - `gradient`: native linear, radial, rectangular, or path type plus the bounded options exposed by Office.js.
   - `pattern`: native preset plus foreground/background colors.
   - `picture_or_texture`: a `/home/user/...` VFS image path plus transparency. The existing bounded image reader validates type and size before Base64 is passed to Office.js.
2. `set_master_theme_color`
   - One of Accent1–6, Dark1–2, Light1–2, Hyperlink, or FollowedHyperlink.
3. `set_layout_background_following`
   - Controls `isMasterBackgroundFollowed` and `areBackgroundGraphicsHidden` through explicit booleans.

Colors are normalized to uppercase `#RRGGBB`; enums are exact allowlists; unknown fields are rejected. Each operation targets explicit IDs obtained from `inspect_slide_masters`. The agent may create multiple operations for all masters, still under one proposal.

The existing XML implementation is renamed `edit_slide_master_xml` and remains available only where the host policy explicitly enables reliable package import. It is hidden on Mac. `edit_slide_xml` and `edit_slide_chart` are unchanged.

## Execution and verification

### Prepare

1. Require PowerPointApi 1.10.
2. Resolve every target master/layout ID and reject missing or duplicate-conflicting operations.
3. Read the exact affected properties, including fill-type-specific values.
4. Build a bounded proposal preview containing IDs, property names, and before/after summaries; never include image bytes.
5. Fingerprint the normalized affected-property snapshot.

### Confirm

1. Re-read all affected properties and compare the fingerprint. Any mismatch returns `proposal_stale` before writing.
2. Apply operations in program order inside bounded `PowerPoint.run` batches.
3. After each sync, read back the affected property and verify semantic equality.
4. After all operations, perform a final full affected-property readback.

### Recovery

On failure or cancellation after writing begins, process applied operations in reverse order:

- Re-read the current value.
- If it equals the captured pre-state, recovery for that property is already complete.
- If it equals the transaction’s verified post-state, restore the captured pre-state and verify it.
- If it equals neither, do not overwrite it and return `office_concurrent_change`.
- If restoration or restoration verification fails, return `office_recovery_failed`.
- If the host cannot classify the state, return `office_state_uncertain` with a safe diagnostic location.

Image-background recovery retains the bounded original image representation only when Office.js exposes it. If the original picture/texture fill cannot be reconstructed from native readback, the proposal must be rejected before writing with `office_api_unsupported`; the tool must not start a transaction it cannot reverse.

## Compatibility and fallback

- PowerPointApi 1.10 present: expose `inspect_slide_masters` and native `edit_slide_master` on Mac and Windows.
- PowerPointApi 1.10 absent: do not advertise native `edit_slide_master`; the agent may use existing slide-local tools, but only as a separately proposed action.
- Windows may additionally expose `edit_slide_master_xml` for native-API gaps after its existing package verification policy passes.
- No automatic fallback occurs inside a confirmed proposal because that would change the approved impact and recovery model.

## Components

- `apps/office-addin/src/skills/powerpoint/browser-powerpoint-adapter.ts`
  - Add native master inspection, snapshot, mutation, semantic readback, and recovery primitives.
- `apps/office-addin/src/skills/powerpoint/powerpoint-skill.ts`
  - Add the read tool, v2 schema/parser, proposal orchestration, renamed XML tool, and capability-aware prompt.
- `apps/office-addin/src/agent/host-runtime.ts`
  - Pass the VFS to the PowerPoint skill and expose tools according to PowerPointApi 1.10.
- `apps/office-addin/tests/powerpoint-skill.test.ts`
  - Cover schemas, capability inventory, confirmation, stale state, native operation mapping, semantic verification, reverse recovery, concurrent edits, and unsupported reversible picture states.
- `apps/office-addin/tests/host-runtime.test.ts`
  - Cover Mac/Windows and 1.10 availability inventories.

No Relay, WisWork PC, pairing protocol, or manifest change is required. Release requires a Taskpane deployment and PowerPoint restart/reload.

## Verification

- Write public inventory and native transaction tests first and demonstrate the expected RED failures.
- Run targeted PowerPoint tests to GREEN.
- Run Office add-in typecheck, lint, changed-file formatting, and the full Office test suite.
- UI-test on the connected Mac PowerPoint: inspect masters, reject proposal, solid background, theme color, gradient/pattern when supported, and forced stale-state behavior.

## Design rationale

| Choice                              | Rationale                                                                   | Tradeoff                                                         |
| ----------------------------------- | --------------------------------------------------------------------------- | ---------------------------------------------------------------- |
| Native PowerPointApi 1.10 first     | Avoids nondeterministic master-package import and enables semantic readback | Requires modern Office hosts                                     |
| Declarative v2 contract             | Expresses supported intent instead of accepting arbitrary XML               | Does not cover every OOXML feature                               |
| Separate inspection tool            | Gives the agent stable target IDs and observable capabilities               | Adds one read step                                               |
| Keep advanced XML under a new name  | Prevents a hidden contract change and preserves Windows-only escape hatch   | Two master tools on capable Windows hosts                        |
| No implicit fallback                | Confirmation remains exact and recovery remains auditable                   | Older hosts require a new slide-local proposal                   |
| Reject unrecoverable picture states | Never begins a write that cannot be safely reversed                         | Some existing picture backgrounds cannot be overwritten natively |
