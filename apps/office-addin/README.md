# WisWork Office Agent

An Office.js task-pane MVP for Word, Excel, and PowerPoint. It connects the shared WisWork `AgentLoop` to the current text selection. The Agent can read immediately, but replacement and append operations are only proposals until the user explicitly confirms a before/after preview.

This is an integration scaffold, not a claim that the production Gateway is enabled.

## Configure local development

Create an untracked `apps/office-addin/.env.development.local` with non-secret browser configuration:

```dotenv
VITE_WISWORK_AUTHORIZATION_URL=https://YOUR_AUTH_HOST/oauth/authorize
VITE_WISWORK_TOKEN_URL=https://YOUR_AUTH_HOST/oauth/token
VITE_WISWORK_CALLBACK_URL=https://localhost:3000/oauth/callback
VITE_WISWORK_CLIENT_ID=YOUR_PUBLIC_OFFICE_CLIENT_ID
VITE_WISWORK_ISSUER=https://YOUR_AUTH_HOST
VITE_WISWORK_MESSAGES_URL=https://wisusage.dev.atominnolab.com/v1/messages
```

Do not put client secrets, access tokens, refresh tokens, authorization codes, or PKCE verifiers in environment files. The add-in keeps access and refresh tokens in memory. Only one-time PKCE state and verifier values use `sessionStorage` while a login redirect is in progress.

Missing or invalid values render an unavailable screen and disable login/chat. The messages URL must also match the fixed WisWork provider endpoint enforced by the transport.

## Run and sideload

From the repository root:

```bash
npm install
npm run dev -w @wiswork/office-addin
```

The development server uses trusted local HTTPS on port 3000. The source `apps/office-addin/public/manifest.xml` is explicitly development-only; sideload that file for local work, then open **WisWork Office Agent** from the task pane.

## Deployment build

A build with all validated environment values emits `dist/manifest.xml` using the callback origin for its task pane and icon, plus exact authentication origins in `AppDomains`. It also emits both `dist/oauth/callback.html` and the exact extensionless `dist/oauth/callback`. An unconfigured build deliberately emits no deployable manifest.

The deployment host must serve `/oauth/callback` as `text/html; charset=utf-8`, or rewrite that exact route (including its query string) internally to `/oauth/callback.html`. The Vite development and preview servers already perform this exact rewrite. Never redirect the callback to the `.html` URL: an HTTP redirect can propagate the authorization code in browser or intermediary history.

## Gateway prerequisites

Before end-to-end login can work, the Gateway operator must:

- register the exact callback `https://localhost:3000/oauth/callback` for the public Office client;
- require PKCE S256 and one-time authorization codes;
- permit the local task-pane origin through an explicit CORS allowlist;
- support the documented refresh-token exchange without a browser client secret;
- return safe OAuth errors without upstream bodies or credentials; and
- expose the fixed WisUsage streaming messages contract expected by `@wiswork/ai-provider`.

## Manual acceptance checklist

Do not mark real Gateway acceptance complete until the prerequisites above are confirmed.

1. Start with no configuration and verify the task pane shows **Agent unavailable**, with no prompt or login controls.
2. Configure the registered development client, sideload the manifest, and verify login returns to `/oauth/callback`, then cleans the visible URL to `/taskpane.html`.
3. Select text and ask the Agent to summarize or improve it. Verify streamed text and tool activity appear and **Stop** cancels an active run.
4. Ask for a replacement. Verify the document does not change when the proposal appears, and the preview clearly shows Before and After.
5. Click **Reject** and verify the document remains unchanged.
6. Request another edit, change the Office selection before confirmation, and verify confirmation reports `proposal_stale` without writing.
7. Request again and click **Confirm change** without changing the selection; verify exactly one replacement or append occurs.
8. Start a new instruction while a proposal is pending and verify the old proposal disappears. Log out and verify conversation and pending proposal are cleared.
9. Inspect browser storage and logs: tokens must not be persisted or printed; only in-progress PKCE state may briefly exist in session storage.

## Checks

```bash
npm run test -w @wiswork/office-addin
npm run typecheck -w @wiswork/office-addin
npm run build -w @wiswork/office-addin
```
