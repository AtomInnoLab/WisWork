# Office Add-in Agent MVP Design

**Date:** 2026-08-20  
**Status:** Approved in conversation; persisted for implementation review

## Goal

Connect the WisWork agent runtime to the Office task-pane add-in so a signed-in user can chat with an agent that reads the current Office selection and proposes simple document edits.

The MVP supports Word, Excel, and PowerPoint through the shared Office document API. It deliberately excludes research-specific retrieval, citations, data analysis, cross-application memory, and autonomous bulk editing.

## Product boundary

The agent provides three tools:

1. `read_selection`: read the current selection as text. This is read-only and may run automatically.
2. `propose_replace_selection`: prepare replacement text and show a before/after preview. It must not mutate the Office document.
3. `propose_append_text`: prepare text to append after the current selection and show a preview. It must not mutate the Office document.

The user must explicitly confirm every proposed write. Confirmation executes the already-previewed value through Office.js. Changed input invalidates the prior proposal. The MVP does not expose delete, format, workbook-wide, document-wide, slide-wide, external navigation, file upload, or connector tools.

## Architecture

The Office renderer reuses `@wiswork/agent-core` for the AgentLoop and implements an Office-specific `AgentSkill`. A browser transport sends normalized agent stream requests to the fixed WisUsage-compatible messages endpoint. The renderer never accepts a model endpoint, model name, serving region, or authorization header from a prompt or tool call.

Browser OAuth uses Authorization Code with PKCE S256. Development configuration uses an HTTPS callback at `https://localhost:3000/oauth/callback`; production must supply an approved HTTPS callback. The Gateway must explicitly register the callback and must validate and consume the PKCE verifier before real login is considered supported.

## OAuth and token handling

- Authorization URL, callback URL, token exchange URL, client ID, issuer, and messages endpoint are non-secret build configuration.
- The PKCE verifier and OAuth state live only in `sessionStorage` and are removed after callback processing or logout.
- Access and refresh tokens live only in memory for the MVP. Reloading the task pane signs the user out.
- Tokens are never placed in local storage, IndexedDB, URLs, logs, error messages, tool output, React state intended for persistence, or Office document content.
- OAuth callback processing requires one `code`, one matching `state`, the expected issuer when supplied, and the exact configured callback URL.
- A 401 may refresh once and retry once. Any second 401 signs the user out.
- Logout clears in-memory tokens, PKCE state, agent history, and pending write proposals.

## Model transport

- The renderer sends only the fixed model request shape used by WisUsage.
- The request includes the fixed serving region required by the existing provider contract.
- Streaming text, tool calls, stop reasons, and safe error codes are translated into the `AgentTransport` interface.
- Cancellation uses `AbortController` and must still complete the AgentLoop callback lifecycle.
- User-visible failures contain only a stable stage/code and HTTP status. They must not include tokens or upstream response bodies.
- The Office task pane CSP and manifest allow only the configured WisWork authorization and messages origins in production builds.

## Interaction states

1. Signed out: explanation and sign-in action.
2. Authorizing: disabled chat with cancellable status.
3. Ready: host status, conversation, and prompt input.
4. Agent working: streaming response, tool activity, and stop action.
5. Proposal pending: before/after preview with Confirm and Reject.
6. Applying: controls disabled while Office.js mutates the selection.
7. Applied: audit entry describing the confirmed operation.
8. Blocked/error: safe recovery action without leaking upstream details.

Only one write proposal may be pending. Starting a new user turn rejects the prior unconfirmed proposal.

## Agent prompt and context

Each turn receives the active host and the current selection text, capped before entering model context. The system prompt tells the model to use tools for document state, never claim an edit happened before confirmation, and avoid destructive or unsupported actions.

Selection text and proposed text have explicit length limits. Tool arguments are schema validated and rejected when fields are missing, extra, or oversized.

## Trust and autonomy

| Action                                                  | Autonomy              | Rationale                                                       |
| ------------------------------------------------------- | --------------------- | --------------------------------------------------------------- |
| Read host and selection                                 | Automatic             | Local, read-only, and visible in the active document            |
| Explain or draft text                                   | Automatic             | Does not mutate the document                                    |
| Replace or append selection                             | Explicit confirmation | User-visible document mutation; preview is cheap and reversible |
| Bulk, structural, formatting, file, or external actions | Refused               | Outside the MVP's bounded and tested action surface             |

The UI keeps a session-only audit trail of tool calls and confirmed writes. Office-native Undo remains available after a write; the add-in does not claim cross-host transactional rollback.

## Failure handling

- Missing Gateway configuration keeps the add-in in a clearly labeled unavailable state.
- Unsupported host or selection coercion returns a safe tool error to the model and user.
- A changed or missing selection at confirmation requires a refreshed preview rather than applying stale text.
- Network interruption leaves the document untouched and preserves the user's prompt for retry.
- Invalid model tool calls are returned to AgentLoop as tool errors; they never execute.
- Prompt injection in document content cannot expand the tool set, change endpoints, bypass confirmation, or access tokens.

## Files and interfaces

- `apps/office-addin/src/auth/*`: browser PKCE session and callback processing.
- `apps/office-addin/src/agent/transport.ts`: fixed WisUsage streaming transport.
- `apps/office-addin/src/agent/office-skill.ts`: bounded Office tool definitions and proposal creation.
- `apps/office-addin/src/agent/use-office-agent.ts`: AgentLoop/UI state coordination.
- `apps/office-addin/src/App.tsx`: signed-out, chat, activity, and confirmation UI.
- `apps/office-addin/src/office-document.ts`: selection reads and confirmed writes.
- `apps/office-addin/src/config.ts`: validated non-secret build configuration.
- `apps/office-addin/tests/*`: auth, transport, skill, confirmation, and existing Office adapter tests.

`@wiswork/agent-core` is consumed as a workspace dependency. `@wiswork/ai-provider` may be reused only if its browser path preserves the fixed-endpoint and safe-error constraints; Electron IPC and main-process auth code are not imported into the Office renderer.

## Verification

- Unit tests cover PKCE/state validation, callback replay rejection, token clearing, refresh-once behavior, cancellation, safe errors, tool schemas, proposal invalidation, and confirm-before-write.
- Existing Office adapter tests remain green.
- TypeScript, ESLint, formatting, Vite build, manifest structural checks, and CSP origin checks pass.
- A manual localhost acceptance run verifies sign-in redirect, callback, streaming chat, read selection, rejected proposal, confirmed replacement, stop, logout, and reload sign-out.
- A real integration build is not accepted until the Gateway owner confirms callback registration, PKCE verifier validation, one-time authorization code consumption, allowed CORS origins, refresh behavior, and safe error responses.

## Rollback

The feature is isolated to `apps/office-addin` and guarded by valid Gateway configuration. Removing the Agent UI/auth files and workspace dependencies restores the existing selection-only lab. No persistent user data, migration, or Office document schema change is introduced.

## Success criteria

- A configured user can sign in and receive a streamed agent response in the Office task pane.
- The agent can read the active selection and prepare replace/append proposals.
- No Office mutation occurs without a fresh, explicit confirmation.
- Tests demonstrate that credentials and upstream response bodies cannot reach logs or user-visible errors.
- Unconfigured or unapproved Gateway environments fail closed.
