# WisWork Office Agent

An Office.js task pane for Word, Excel, and PowerPoint. It reuses the signed-in WisWork PC account
through a consented cloud Relay. Wispaper credentials remain in the PC process; the task pane keeps
only a short-lived socket-bound capability in memory. Document writes still require confirmation.

## Shipped tool surface

After `Office.onReady()`, the task pane composes the shared browser skill with exactly one Word,
Excel, or PowerPoint skill. Unsupported hosts fail closed.

- Shared: bounded VFS `read`, sandboxed `bash` (`pwd`, `ls`, and `cat`), session-file
  upload, and strict single-file `SKILL.md` installation. Installed skill metadata is added to the
  Agent context dynamically; package folders and auxiliary files are not yet exposed in the UI.
- Word: `get_document_text`, `get_document_structure`, `get_ooxml`, `screenshot_document`,
  `execute_office_js`.
- Excel: `get_cell_ranges`, `get_range_as_csv`, `search_data`, `screenshot_range`,
  `get_all_objects`, `set_cell_range`, `clear_cell_range`, `copy_to`,
  `modify_sheet_structure`, `modify_workbook_structure`, `resize_range`, `modify_object`,
  `eval_officejs`.
- PowerPoint: `screenshot_slide`, `list_slide_shapes`, `read_slide_text`, `verify_slides`,
  `edit_slide_text`, `duplicate_slide`, `edit_slide_xml`, `edit_slide_chart`,
  `edit_slide_master`, `execute_office_js`.

Word screenshot exports a bounded PDF through Office.js and renders a bounded page to a
model-visible PNG. Word `execute_office_js` accepts only a JSON declarative program (version 1,
maximum 32 allowlisted operations); it never evaluates JavaScript. PowerPoint package operations
are tracked separately. File conversions remain blocked until they run in a terminateable worker
with ZIP metadata, decompression, cell, pixel, and output quotas. Web retrieval remains blocked
because there is no fixed authenticated PC bridge route, so neither capability is advertised.
Browser `bash` is not a
native shell and has no PC filesystem, process, socket, credential, or package-install access.
Supported mutations show a structured title, impact, before/after data, preview, and code when
present, and execute only after explicit confirmation. Logout or bridge loss disposes proposals,
uploaded VFS files, and installed-skill session state.

## Connect WisWork PC

The task pane connects by default to the fixed secure Relay at
`wss://office.8-216-134-194.sslip.io/office-relay`. Office and WisWork PC both make outbound WSS
connections; no inbound port is opened on the user's computer. Enter the six-digit Office code in
the signed-in PC app and approve the matching host/code. Credentials remain in WisWork PC.

## Develop and sideload

From the repository root:

```bash
npm install
npm run dev:office
```

The development server uses trusted local HTTPS on port 3000. Sideload
`apps/office-addin/public/manifest.xml`, open **WisWork Office Agent**, enter its code in WisWork PC,
then approve the matching request.

## Deployment build

Configure only the deployment origin for the default Relay build:

```bash
VITE_WISWORK_ADDIN_ORIGIN=https://office.example \
npm run build -w @wiswork/office-addin
```

A valid configured build emits `dist/manifest.xml`; an invalid build emits no deployable manifest.
The task-pane CSP allows only the fixed WSS Relay. There are no OAuth callback pages, direct
WisUsage connections, wildcard origins, or source maps in the deployment output.

The local HTTP bridge is rollback-only. It is never selected automatically. A coordinated rollback
build must set `VITE_WISWORK_OFFICE_TRANSPORT=loopback` and configure the same bounded port list on
Office and PC. Remove that flag to return to Relay mode.

The new host/shared registries are enabled by default. Build with
`VITE_WISWORK_OFFICE_HOST_SKILLS=0` to roll back to the legacy selection-only skill without
changing the PC bridge, identity, manifest, or stored user data.

## Operational behavior

- **PC offline:** Office keeps the short-lived code visible so it can be entered after PC starts.
- **PC signed out:** approval is refused; sign in through the existing WisWork PC flow first.
- **Approval:** every new task-pane session requires visible approval in WisWork PC. Approve only when the same six-digit verification code is visible in both Office and the PC confirmation dialog.
- **Revocation:** Office logout/disconnect drops the in-memory capability. PC logout, bridge shutdown, or PC restart revokes every pairing and active stream.
- **Relay restart/network loss:** the in-memory session is revoked; reconnect creates a new pairing.
- **Rollback:** deploy Office and PC rollback settings together. There is no silent HTTP downgrade.

## Manual Office acceptance

Relay behavior must be checked on Windows and macOS desktop Office and Word Web before release:

1. Start WisWork PC signed in; sideload the configured manifest in Word, Excel, and PowerPoint.
2. Confirm the task pane opens only the fixed WSS Relay and needs no loopback/PNA exception.
3. Connect, approve in PC, and verify the Agent conversation appears and streams using the same Wispaper account and credits as PC.
4. Verify Reject never changes the document, stale proposals fail, and Confirm applies exactly one replacement or append.
5. Log out of WisWork PC during an active stream. Verify the stream stops, Office clears conversation/proposals, and reconnect requires a new approval.
6. Stop WisWork PC and verify Office returns to its offline state without retaining a capability. Restart and pair again.
7. Restart the Relay and verify both clients revoke in-memory state and require a new pairing.
8. Repeat pairing and one streamed request in Word Web.
9. Inspect Office storage, logs, and network frames: no Wispaper credential may appear.

For the host-tool release candidate, also complete these checks on both Windows and macOS desktop
Office (with host-specific API-set acceptance still required):

1. Verify each host advertises only shared `read`/`bash` plus its exact inventory above, with no
   cross-host tools.
2. Run one supported read and screenshot; if the required API is unavailable, verify the stable
   unsupported result rather than a success claim.
3. Create a supported Excel mutation and PowerPoint text/duplicate mutation. Verify title, impact,
   targets, before/after or preview, Reject, stale-state rejection, and exactly-once Confirm.
4. Exercise every release-blocked tool and verify its documented error and zero mutation.
5. Upload and read a session file, install a valid file named `SKILL.md`, verify its metadata enters
   the next Agent request, then verify traversal and native-shell/network syntax are denied.
6. Log out during an active stream and pending confirmation; verify conversation, proposal, VFS,
   and installed-skill state are cleared before reconnecting.

## Browser conversion dependency audit

The Word screenshot runtime uses the exactly pinned PDF.js 5.7.284 package (Apache-2.0). PDF input
comes only from Office's bounded document export; loading disables worker fetch and WebAssembly,
sets image/page/output bounds, and uses a fixed bundled worker URL. JSZip 3.10.1
(`MIT OR GPL-3.0-or-later`) and
fast-xml-parser 5.10.1 (MIT) are pinned for the separately reviewed PowerPoint package runtime, not
as generally available converters. The configured production build emits no source map. Any
version change requires repeating license, CSP, bundle-size, and vulnerability review.

## Checks

```bash
npm run test -w @wiswork/office-bridge
npm run test -w @wiswork/office-addin
npm run typecheck -w @wiswork/office-bridge
npm run typecheck -w @wiswork/office-addin
VITE_WISWORK_ADDIN_ORIGIN=https://office.example npm run build -w @wiswork/office-addin
```
