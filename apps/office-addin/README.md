# WisWork Office Agent

An Office.js task pane for Word, Excel, and PowerPoint. It reuses the signed-in WisWork PC account and Agent transport through a consented local bridge. Wispaper credentials remain in the PC process; the task pane keeps only a short-lived bridge capability in memory. Document writes still require an explicit before/after confirmation.

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

## Checks

```bash
npm run test -w @wiswork/office-bridge
npm run test -w @wiswork/office-addin
npm run typecheck -w @wiswork/office-bridge
npm run typecheck -w @wiswork/office-addin
VITE_WISWORK_ADDIN_ORIGIN=https://office.example npm run build -w @wiswork/office-addin
```
