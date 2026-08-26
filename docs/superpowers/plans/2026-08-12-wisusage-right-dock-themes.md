# WisUsage, Right AI Dock, and Themes Implementation Plan

## Goal and non-goals

Deliver authenticated WisUsage model calls, right-side AI docks across all document applications, and persistent Light/Dark UI themes. Do not change document file formats, authored page colors, optional third-party AI providers, or add usage billing logic in the client.

## Global constraints

- No access token crosses into renderer-visible state or logs.
- WisUsage endpoint/model are authoritative main-process constants.
- Exactly one AI panel instance remains mounted per document.
- Theme state is shell-owned, validated, persisted, and synchronized to all views.
- All runtime changes follow RED/GREEN TDD.

## Deliverable 1: authenticated WisUsage transport

Files: `packages/ai-provider/src/{providers,main-config,ipc,stream,chat}.ts`, provider tests, shell AI registration and auth integration.

Acceptance: a logged-in session streams from `/v1/messages` with `Authorization: Bearer <current access token>`, Anthropic request shape, fixed Qwen model, and refresh-aware retry. Logged-out sessions fail before network I/O. No `WISWORK_MODEL_API_KEY` is required or exposed.

Sequence: add failing main-config/stream/IPC tests; introduce async token-backed request resolution and WisUsage transport; wire shell auth client; run provider/auth/shell suites and commit.

## Deliverable 2: universal right AI dock

Files: document renderer Apps/styles and focused layout tests for Docs, Sheets, Slides, PDF, and LaTeX.

Acceptance: expanded and collapsed AI docks render at the right edge in normal workspace flows, retain state, and preserve resizing/collapse behavior.

Sequence: add failing source/render assertions; reorder workspace children and normalize flex order/resizer edge; run renderer suites/builds and commit.

## Deliverable 3: Light/Dark themes

Files: shell shared API/main/preload/home UI, `packages/ui` theme utilities/styles, renderer entrypoints/styles, tests.

Acceptance: home exposes Light/Dark control, selection persists across restart, Electron applies it to every view, application chrome visibly changes, and document authored colors remain stable.

Sequence: add failing persistence/IPC/UI/style tests; implement main-owned preference and switch; add common semantic dark tokens plus module overrides; run module tests/builds and commit.

## Release verification

Run all affected tests, full monorepo typecheck, lint, format check, production builds, E2E, `git diff --check`, and a targeted token/key scan. Independently review security boundaries and the aggregate diff before pushing the existing integration branch.
