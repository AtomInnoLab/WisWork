# WisWork Office Agent

An Office.js task pane for Word, Excel, and PowerPoint. It reuses the signed-in WisWork PC account and Agent transport through a consented local bridge. Wispaper credentials remain in the PC process; the task pane keeps only a short-lived bridge capability in memory. Document writes still require an explicit before/after confirmation.

## Shipped tool surface

After `Office.onReady()`, the task pane composes the shared browser skill with exactly one Word,
Excel, or PowerPoint skill. Unsupported hosts fail closed.

- Shared: bounded VFS `read`, sandboxed `bash` (`pwd`, `ls`, `cat`, `pdf-to-text`,
  `pdf-to-images`, `docx-to-text`, and `xlsx-to-csv`), session-file
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
are tracked separately. Web retrieval remains blocked because there is no fixed authenticated PC
bridge route, so no web tool is advertised. Browser `bash` is not a
native shell and has no PC filesystem, process, socket, credential, or package-install access.
Supported mutations show a structured title, impact, before/after data, preview, and code when
present, and execute only after explicit confirmation. Logout or bridge loss disposes proposals,
uploaded VFS files, and installed-skill session state.

## Start WisWork PC

The Office bridge is disabled by default. Start the PC app with an exact HTTPS add-in origin:

```bash
WISWORK_OFFICE_BRIDGE_ENABLED=1 \
WISWORK_OFFICE_ORIGIN=https://office.example \
npm run dev -w @wiswork/shell
```

`WISWORK_OFFICE_ORIGIN` must exactly match the deployed task-pane origin. By default, the PC app tries a 64-port pool (`43127` first, then the remaining ports in `43120–43183`) and binds the first available port. The bridge binds only `127.0.0.1` and never exposes PC access or refresh tokens.

To use a custom pool, configure the same ordered, comma-separated list on PC and in the add-in build. Lists must contain 1–128 unique decimal ports from 1 through 65535:

```bash
WISWORK_OFFICE_BRIDGE_PORTS=43127,43128,43129
```

## Develop and sideload

From the repository root:

```bash
npm install
npm run dev:office
```

The development server uses trusted local HTTPS on port 3000. Start the PC bridge with `WISWORK_OFFICE_ORIGIN=https://localhost:3000`, sideload `apps/office-addin/public/manifest.xml`, and open **WisWork Office Agent**. Click **Connect to WisWork PC**, then approve the request in the PC app.

## Deployment build

Configure the deployment origin and, when needed, the bridge port pool shared with WisWork PC:

```bash
VITE_WISWORK_ADDIN_ORIGIN=https://office.example \
VITE_WISWORK_PC_BRIDGE_PORTS=43127,43128,43129 \
npm run build -w @wiswork/office-addin
```

A valid configured build emits `dist/manifest.xml`. An unconfigured or invalid build emits no deployable manifest. If omitted, `VITE_WISWORK_PC_BRIDGE_PORTS` uses the default 64-port pool. A custom list must exactly match PC runtime `WISWORK_OFFICE_BRIDGE_PORTS`. The generated manifest contains only the exact add-in HTTPS origin; task-pane CSP explicitly enumerates the numeric loopback endpoints in the pool. There are no OAuth callback pages, auth domains, direct WisUsage connection, wildcard origins, or source maps in the deployment output.

The new host/shared registries are enabled by default. Build with
`VITE_WISWORK_OFFICE_HOST_SKILLS=0` to roll back to the legacy selection-only skill without
changing the PC bridge, identity, manifest, or stored user data.

## Operational behavior

- **PC offline:** Office shows **Open WisWork PC** and can retry after the app starts.
- **PC signed out:** approval is refused; sign in through the existing WisWork PC flow first.
- **Approval:** every new task-pane session requires visible approval in WisWork PC. Approve only when the same six-digit verification code is visible in both Office and the PC confirmation dialog.
- **Revocation:** Office logout/disconnect drops the in-memory capability. PC logout, bridge shutdown, or PC restart revokes every pairing and active stream.
- **Port conflict:** the PC app tries the next configured loopback port only when a port is already occupied. It never falls back to another address or a public bind. Startup fails if the whole pool is occupied or a non-conflict bind error occurs.
- **Diagnostics:** the trusted WisWork account menu reports the local bridge as `disabled`, `ready:<port>`, or `error`; errors do not expose network or authentication details.
- **Rollback:** unset `WISWORK_OFFICE_BRIDGE_ENABLED` (or set it to `0`) and deploy the prior add-in build. Restarting WisWork PC clears all in-memory grants; no data migration is required.

## Manual Windows/macOS acceptance

Private Network Access and Office WebView behavior must be checked on both supported desktop platforms before release:

1. Start WisWork PC signed in with the bridge environment above; sideload the configured manifest in Word, Excel, and PowerPoint.
2. Confirm the HTTPS task pane discovers the selected endpoint through the bounded `/v1/office/health` probe and can preflight/fetch it, including `Access-Control-Allow-Private-Network: true` where the WebView requests it.
3. Connect, approve in PC, and verify the Agent conversation appears and streams using the same Wispaper account and credits as PC.
4. Verify Reject never changes the document, stale proposals fail, and Confirm applies exactly one replacement or append.
5. Log out of WisWork PC during an active stream. Verify the stream stops, Office clears conversation/proposals, and reconnect requires a new approval.
6. Stop WisWork PC and verify Office returns to its offline state without retaining a capability. Restart and reconnect.
7. Occupy `43127`; verify PC selects the next free configured port and Office still connects. Then occupy the entire configured pool and verify startup fails without listening publicly.
8. Inspect Office storage, logs, and network responses: no Wispaper access token, refresh token, authorization code, or upstream secret may appear.

For the host-tool release candidate, also complete these checks on both Windows and macOS desktop
Office (Office Web remains blocked pending PNA and API-set acceptance):

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

The conversion/screenshot runtime uses already repository-pinned browser libraries: PDF.js
5.7.284 (Apache-2.0), JSZip 3.10.1 (MIT), and fast-xml-parser 5.3.4 (MIT). Inputs remain inside the
session VFS; PDF loading disables worker fetch and WebAssembly, sets parser/image/page/output
bounds, and uses a fixed bundled worker URL. The configured production build emits no source map;
its expected large artifacts are the PDF worker (about 1.29 MB minified) and lazy PDF renderer
chunk. Any version change requires repeating license, CSP, bundle-size, and vulnerability review.

## Checks

```bash
npm run test -w @wiswork/office-bridge
npm run test -w @wiswork/office-addin
npm run typecheck -w @wiswork/office-bridge
npm run typecheck -w @wiswork/office-addin
VITE_WISWORK_ADDIN_ORIGIN=https://office.example npm run build -w @wiswork/office-addin
```
