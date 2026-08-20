# WisWork Office Agent

An Office.js task pane for Word, Excel, and PowerPoint. It reuses the signed-in WisWork PC account and Agent transport through a consented local bridge. Wispaper credentials remain in the PC process; the task pane keeps only a short-lived bridge capability in memory. Document writes still require an explicit before/after confirmation.

## Start WisWork PC

The Office bridge is disabled by default. Start the PC app with an exact HTTPS add-in origin:

```bash
WISWORK_OFFICE_BRIDGE_ENABLED=1 \
WISWORK_OFFICE_ORIGIN=https://office.example \
WISWORK_OFFICE_BRIDGE_PORT=43127 \
npm run dev -w @wiswork/shell
```

`WISWORK_OFFICE_ORIGIN` must exactly match the deployed task-pane origin. The port defaults to `43127`; changing it requires setting the same port in the PC runtime and add-in build. The bridge binds only `127.0.0.1` and never exposes PC access or refresh tokens.

## Develop and sideload

From the repository root:

```bash
npm install
npm run dev:office
```

The development server uses trusted local HTTPS on port 3000. Start the PC bridge with `WISWORK_OFFICE_ORIGIN=https://localhost:3000`, sideload `apps/office-addin/public/manifest.xml`, and open **WisWork Office Agent**. Click **Connect to WisWork PC**, then approve the request in the PC app.

## Deployment build

Configure the deployment origin and the bridge port shared with WisWork PC:

```bash
VITE_WISWORK_ADDIN_ORIGIN=https://office.example \
VITE_WISWORK_PC_BRIDGE_PORT=43127 \
npm run build -w @wiswork/office-addin
```

A valid configured build emits `dist/manifest.xml`. An unconfigured or invalid build emits no deployable manifest. `VITE_WISWORK_PC_BRIDGE_PORT` defaults to `43127` and accepts only an integer from 1 through 65535. It must equal PC runtime `WISWORK_OFFICE_BRIDGE_PORT`. The generated manifest contains only the exact add-in HTTPS origin; task-pane CSP permits connections only to itself and the numeric loopback endpoint at that port. There are no OAuth callback pages, auth domains, direct WisUsage connection, wildcard origins, or source maps in the deployment output.

## Operational behavior

- **PC offline:** Office shows **Open WisWork PC** and can retry after the app starts.
- **PC signed out:** approval is refused; sign in through the existing WisWork PC flow first.
- **Approval:** every new task-pane session requires visible approval in WisWork PC. Approve only when the same six-digit verification code is visible in both Office and the PC confirmation dialog.
- **Revocation:** Office logout/disconnect drops the in-memory capability. PC logout, bridge shutdown, or PC restart revokes every pairing and active stream.
- **Port conflict:** the PC app reports bridge startup failure and does not fall back to another address or public bind. Stop the conflicting process or deliberately configure the same new port in both PC and a rebuilt add-in.
- **Diagnostics:** the trusted WisWork account menu reports the local bridge as `disabled`, `ready`, or `error`; `error` includes port conflicts without exposing network or authentication details.
- **Rollback:** unset `WISWORK_OFFICE_BRIDGE_ENABLED` (or set it to `0`) and deploy the prior add-in build. Restarting WisWork PC clears all in-memory grants; no data migration is required.

## Manual Windows/macOS acceptance

Private Network Access and Office WebView behavior must be checked on both supported desktop platforms before release:

1. Start WisWork PC signed in with the bridge environment above; sideload the configured manifest in Word, Excel, and PowerPoint.
2. Confirm the HTTPS task pane can preflight and fetch `http://127.0.0.1:43127`, including `Access-Control-Allow-Private-Network: true` where the WebView requests it.
3. Connect, approve in PC, and verify the Agent conversation appears and streams using the same Wispaper account and credits as PC.
4. Verify Reject never changes the document, stale proposals fail, and Confirm applies exactly one replacement or append.
5. Log out of WisWork PC during an active stream. Verify the stream stops, Office clears conversation/proposals, and reconnect requires a new approval.
6. Stop WisWork PC and verify Office returns to its offline state without retaining a capability. Restart and reconnect.
7. Occupy port `43127`; verify PC reports the conflict and never listens publicly. Release the port and restart successfully.
8. Inspect Office storage, logs, and network responses: no Wispaper access token, refresh token, authorization code, or upstream secret may appear.

## Checks

```bash
npm run test -w @wiswork/office-bridge
npm run test -w @wiswork/office-addin
npm run typecheck -w @wiswork/office-bridge
npm run typecheck -w @wiswork/office-addin
VITE_WISWORK_ADDIN_ORIGIN=https://office.example npm run build -w @wiswork/office-addin
```
