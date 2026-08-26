# WisUsage, Right AI Dock, and Application Themes

## Goal

Use the authenticated WisWork session token for all managed model requests, place every document AI assistant on the right, and add persistent Light/Dark application themes.

## Boundaries

- The managed model endpoint is exactly `https://wisusage.dev.atominnolab.com/v1/messages` and the model is exactly `qwen/qwen3.8-max`.
- The access token is resolved and refreshed only in the shell main process. It is never persisted in AI settings, returned over IPC, placed in a renderer URL, or logged.
- Existing optional non-WisWork providers remain unchanged. Managed WisWork requests use the Anthropic Messages streaming protocol exposed by WisUsage.
- AI docks in Docs, Sheets, Slides, PDF, and LaTeX are the final/rightmost workspace child. Collapsed rails remain on the right.
- Theme selection has exactly two values, `light` and `dark`, is persisted by the shell, and is applied through Electron's native theme to the shell and all document WebContents.
- Theme changes affect application chrome and workspaces. Document/page/slide authored colors remain unchanged.

## Architecture

`@wiswork/auth` continues to own token refresh. `@wiswork/ai-provider` receives a main-process-only async token resolver and builds the authoritative WisUsage provider configuration; the renderer-provided managed configuration remains ignored. WisUsage streaming is parsed through the existing Anthropic Messages adapter with a fixed endpoint and without an Anthropic API-key header.

The shell owns a persisted theme preference and applies it to Electron `nativeTheme.themeSource`. The home renderer exposes the Light/Dark switch; all module styles respond to `prefers-color-scheme`, while shared tests enforce the contract. AI layout is normalized by DOM order and flex ordering, not by duplicating panels.

## Failure Handling and Security

- Logged out or refresh failure returns `auth_required`; no request is sent.
- HTTP 401 is allowed through the authenticated request boundary so the auth client can refresh once and retry with the new token.
- Upstream failures retain stable, non-secret error codes. Response bodies and bearer tokens are not logged.
- Invalid theme IPC payloads are rejected and untrusted senders cannot read or mutate theme settings.

## Verification

- RED/GREEN provider tests assert exact URL, bearer token, request shape, streaming deltas/tool calls, refresh retry, and absence of the old service-key dependency.
- Layout tests assert AI dock ordering for all five document modules.
- Theme tests assert persistence, validation, broadcast/application, home switch behavior, and dark-style coverage for every renderer.
- Run affected package suites, monorepo typecheck/lint/format, production builds, Electron E2E where available, and credential scans.

## Rollback

Each capability is isolated in its own commit. Reverting the WisUsage commit restores the previous managed transport; reverting layout or theme commits does not affect authentication or document data.
