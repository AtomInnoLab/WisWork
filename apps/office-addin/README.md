# WisWork Office Agent

An Office.js task pane for Word, Excel, and PowerPoint. It reuses the signed-in WisWork PC account
through a consented cloud Relay. Wispaper credentials remain in the PC process; the task pane keeps
only a short-lived socket-bound capability in memory. Document writes still require confirmation.

## Shipped tool surface

After `Office.onReady()`, the task pane composes the shared browser skill with exactly one Word,
Excel, or PowerPoint skill. Unsupported hosts fail closed.

- Shared: bounded VFS `read`, sandboxed `bash` (`cat`, `pdf-to-text`,
  `pdf-to-images`, `docx-to-text`, and `xlsx-to-csv`), session-file
  upload, and bounded ZIP skill-package installation/removal. A package contains exactly one root
  `SKILL.md` plus allowlisted UTF-8 text, PNG, JPEG, or WebP resources. Installed skill metadata is
  added to Agent context dynamically and package resources are mounted read-only in the session VFS.
- Word: `get_document_text`, `get_document_structure`, `get_ooxml`, `screenshot_document`,
  `write_document`, `execute_office_js`.
- Excel: `get_cell_ranges`, `get_range_as_csv`, `search_data`, `screenshot_range`,
  `get_all_objects`, `set_cell_range`, `clear_cell_range`, `copy_to`,
  `modify_sheet_structure`, `modify_workbook_structure`, `resize_range`, `modify_object`,
  `eval_officejs`; when ExcelApi 1.9 is present, also `csv-to-sheet`, `sheet-to-csv`, and
  `image-to-sheet`.
- PowerPoint: `screenshot_slide`, `list_slide_shapes`, `read_slide_text`, `verify_slides`,
  `edit_slide_text`, `duplicate_slide`, `edit_slide_xml`, `edit_slide_chart`,
  `edit_slide_master`, `execute_office_js`; when PowerPointApi 1.8 is present, also
  `insert-image`.

CSV import/export is capped at 500 rows, 100 columns, 10,000 cells, and 2 MiB. Export prefixes
formula-like cells before quoting so opening the CSV cannot execute formula payloads. Image tools
accept only magic-validated PNG/JPEG bytes from the session VFS, capped at 2 MiB, 8,192 pixels per
dimension and 16,777,216 decoded pixels. Imports are immutable proposals: Confirm rechecks the
range/slide fingerprint, dispatches once, and verifies the exact cells or created image ID and
geometry. CSV writes retain a bounded formulas/value snapshot; image writes retain the pre-write
shape IDs. A partial write, cancellation, or semantic mismatch restores the range or deletes the
inserted shape and proves recovery before returning; an unprovable recovery is terminal
`office_recovery_failed`. Word advertises no import/media command because the approved inventory has no
corresponding Word tool contract. Older Office API sets keep these tools unadvertised rather than
simulating support.

Word screenshot exports a bounded PDF through Office.js and renders a bounded page to a
model-visible PNG. Word `execute_office_js` accepts only a JSON declarative program (version 1,
maximum 32 allowlisted operations); it never evaluates JavaScript. PowerPoint package operations
are tracked separately. File conversions run only in the bounded terminateable worker described
below. Web retrieval remains independently capability-negotiated through the authenticated PC
Relay route. Browser `bash` is not a
native shell and has no PC filesystem, process, socket, credential, or package-install access.
For ordinary drafting, Word uses the structured `write_document` tool with bounded text and an
exact `replace`, `append`, or `prepend` mode. It creates one confirmation card and reports
`awaiting_user_confirmation` as success; repeated writes while that card is pending reuse it rather
than replacing the proposal or retrying the document mutation.
Supported mutations show a structured title, impact, and bounded before/after comparison or
readable preview; declarative code remains an internal protocol and is never shown as approval UI.
They execute only after explicit confirmation. Conversion commands accept one matching
file from the session VFS and mount outputs atomically only after a dedicated Web Worker completes.
DOCX/XLSX archives are rejected before inflate for invalid central/local metadata, duplicate or
case-colliding normalized paths, traversal, encryption, special file types, unsupported methods,
or claimed quotas. Selected entries are then inflated incrementally; the stream is paused and the
worker fails before retaining bytes beyond the per-entry or aggregate actual-output limits. PDF
page/text/image limits are enforced in the same terminateable worker. `pdf-to-images` fails with
`conversion_unsupported` when the Office WebView has no `OffscreenCanvas`; it never falls back to
main-thread rendering. Logout or bridge loss disposes proposals,
uploaded VFS files, and installed-skill session state.

Skill packages are parsed in a terminateable worker. Compressed input, entry count, normalized path,
per-file and aggregate inflated bytes are hard bounded; inflate is streamed and stops at the first
limit breach. Raw central/local headers, names, flags, methods, offsets, ranges, data descriptors,
sizes and CRCs must agree, and every ZIP64 form is refused. Traversal, Unicode/case collisions,
links, executable/special entries, invalid UTF-8,
unsupported formats, and duplicate packages fail atomically. Packages can contribute instructions
and declarative resources only. PNG/JPEG/WebP structure and dimensions are validated under a pixel
cap. The exact bounded `SKILL.md` body enters Agent context; packages cannot register tools, code,
commands, network destinations,
Office APIs, or any additional authority. New task, logout, and disposal remove all package mounts.

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

Capability families have independent exact build-time rollback flags. Set any flag to `0` to
remove that family without rolling back the workspace or Relay identity:

- `VITE_WISWORK_OFFICE_CONVERSIONS=0`
- `VITE_WISWORK_OFFICE_SKILL_PACKAGES=0`
- `VITE_WISWORK_OFFICE_IMPORT_MEDIA=0`

Only `0` and `1` are accepted; invalid values make the deployment fail closed.

Safe remote failure diagnostics are enabled by default for Relay builds. They contain only bounded
host/build/tool/phase identifiers, stable error codes, requirement-set support, and allowlisted
Office error identifiers; prompts, document content, tool inputs, formulas, OOXML, screenshots,
tokens, raw error messages, and stacks are never sent. Use **复制诊断信息** in the task-pane session
menu to copy the bounded local diagnostic ring. Set the exact build flag
`VITE_WISWORK_OFFICE_REMOTE_DIAGNOSTICS=0` to retain local export while rolling remote diagnostics
back; other values invalidate the deployment configuration. Relay operators can correlate a copied
`trace_id` with structured `office_diagnostic` service-log events. Logs should be retained for no
more than seven days by the deployment log policy.

Web tools remain unadvertised because there is no compiled, reviewed retrieval-service
attestation. Their activation change must add its capability flag and Office v2 composition in the
same deployment; this build intentionally exposes no no-op Web flag.

The Agent conversation workspace has an independent fail-closed rollback. Build with the exact
flag `VITE_WISWORK_OFFICE_WORKSPACE=0` to retain the legacy task-pane presentation while leaving
Relay identity, the Agent harness, host tools, and confirmation semantics unchanged. Omit the flag
or set it to `1` for the new workspace; other values invalidate the deployment configuration.

## Operational behavior

The connected add-in uses the same task-oriented interaction hierarchy as WisWork PC: a compact
host/connection header, a bounded multi-turn timeline, observable tool activity, inline proposal
review, and a sticky multiline composer. Enter sends, Shift+Enter inserts a line break, and Stop
cancels the active run. Attachments and installed skills live in temporary management panels and
are cleared with the conversation by **New task**, logout, Relay loss, or disposal. The workspace
is designed for 280–500 px task panes and follows the shared light/dark tokens, forced-colors, and
reduced-motion preferences. Every document mutation still requires an explicit inline confirmation.

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
5. Upload and read a session file; install and remove a valid bounded skill ZIP; verify its metadata
   enters the next Agent request and its assets are read-only. Verify traversal, collisions,
   executable entries, unsupported assets, and native-shell/network syntax are denied.
6. Log out during an active stream and pending confirmation; verify conversation, proposal, VFS,
   and installed-skill state are cleared before reconnecting.

## Browser conversion dependency audit

The Word screenshot runtime uses the exactly pinned PDF.js 5.7.284 package (Apache-2.0). PDF input
comes only from Office's bounded document export; loading disables worker fetch and WebAssembly,
sets image/page/output bounds, and uses a fixed bundled worker URL. The conversion worker also uses
that exact PDF.js version with worker fetch, font loading, and eval support disabled. JSZip 3.10.1
(`MIT OR GPL-3.0-or-later`) is pinned for raw ZIP parsing plus incremental DEFLATE/STORE streams;
WisWork validates raw central and local records before handing the bytes to JSZip and never calls
the eager `entry.async()` conversion path. fast-xml-parser 5.10.1 (MIT) is pinned for validated,
bounded OOXML parts. These packages are also used by the separately reviewed PowerPoint package
runtime. All are already production dependencies covered by `npm run licenses`; this change adds
no native module or remote runtime dependency. The configured production build emits separate
conversion/PDF worker chunks and no source map. Any version change requires repeating license,
CSP, bundle-size, forged-archive, and vulnerability review.

## Checks

```bash
npm run test -w @wiswork/office-bridge
npm run test -w @wiswork/office-addin
npm run typecheck -w @wiswork/office-bridge
npm run typecheck -w @wiswork/office-addin
VITE_WISWORK_ADDIN_ORIGIN=https://office.example npm run build -w @wiswork/office-addin
```
