# Office Agent Tool Parity Design

## Purpose

WisWork Office will reuse the existing `@wiswork/agent-core` harness without changing the PC Agent harness. The task pane will load host-specific Word, Excel, or PowerPoint skills plus shared browser-only skills. The reference capability baseline is the public tool inventory in `hewliyang/office-agents`; WisWork retains its own authentication, transport, confirmation, and security boundaries.

## Goals

- Support the documented Word, Excel, and PowerPoint tool names and their core user-visible semantics.
- Keep the Agent loop, cancellation, tool routing, compaction, and error model in `@wiswork/agent-core`.
- Execute Office document operations only inside the Office task pane through Office.js.
- Reuse the signed-in WisWork PC identity and credits only through the existing bounded messages bridge.
- Provide shared task-pane VFS, file reading, sandboxed commands, web retrieval, and installable `SKILL.md` packages without exposing the PC filesystem or credentials.
- Preserve confirmation-first behavior for every document mutation, including raw Office.js and OOXML/XML edits.

## Non-goals

- Do not modify or migrate the PC Agent harness.
- Do not expose PC skills, Electron IPC, OS commands, PC files, or access/refresh tokens to the task pane.
- Do not add BYOK provider configuration or an independent OAuth flow to Office.
- Do not claim parity merely because a tool name exists: unsupported Office API sets must return a stable `office_api_unsupported` error.
- Do not execute arbitrary native shell commands. `bash` is a browser sandbox over the task-pane VFS only.

## Architecture

`AgentLoop` remains the single harness. At Office startup, a host registry composes shared browser skills with exactly one host skill. The host skill owns Office.js calls and exposes read operations directly. Mutating tools produce immutable proposals containing bounded preview data and a captured execution closure; only the existing task-pane confirmation action may execute one.

```text
@wiswork/agent-core AgentLoop
  ├─ PC-bridge AgentTransport
  └─ composeSkills()
       ├─ shared browser skill: read, bash, conversions, web, SKILL.md
       └─ exactly one host skill
            ├─ Word
            ├─ Excel
            └─ PowerPoint
```

## Tool compatibility baseline

### Word

- `get_document_text`
- `get_document_structure`
- `get_ooxml`
- `screenshot_document`
- `execute_office_js`
- shared `read` and `bash`
- commands: `pdf-to-text`, `pdf-to-images`, `docx-to-text`, `xlsx-to-csv`, `web-search`, `web-fetch`, `image-search`

### Excel

- `get_cell_ranges`
- `get_range_as_csv`
- `search_data`
- `screenshot_range`
- `get_all_objects`
- `set_cell_range`
- `clear_cell_range`
- `copy_to`
- `modify_sheet_structure`
- `modify_workbook_structure`
- `resize_range`
- `modify_object`
- `eval_officejs`
- shared `read` and `bash`
- commands: `csv-to-sheet`, `sheet-to-csv`, `pdf-to-text`, `pdf-to-images`, `docx-to-text`, `xlsx-to-csv`, `image-to-sheet`, `web-search`, `web-fetch`

### PowerPoint

- `screenshot_slide`
- `list_slide_shapes`
- `read_slide_text`
- `verify_slides`
- `execute_office_js`
- `edit_slide_text`
- `edit_slide_xml`
- `edit_slide_chart`
- `edit_slide_master`
- `duplicate_slide`
- shared `read` and `bash`
- commands: `pdf-to-text`, `pdf-to-images`, `docx-to-text`, `xlsx-to-csv`, `insert-image`, `web-search`, `web-fetch`

## Permission and autonomy model

| Capability                            | Policy                                                                                          |
| ------------------------------------- | ----------------------------------------------------------------------------------------------- |
| Read Office state and structure       | Automatic, bounded, cancellable                                                                 |
| Screenshot/verification               | Automatic, bounded output                                                                       |
| Read VFS and deterministic conversion | Automatic within VFS quotas                                                                     |
| Web search/fetch                      | Automatic through an allowlisted WisWork service; source URLs shown                             |
| VFS writes                            | Automatic, reversible, quota-bound                                                              |
| Structured Office mutations           | Immutable preview and explicit confirmation                                                     |
| Workbook/sheet/master/XML mutations   | Impact summary, preview, explicit confirmation, pre-write snapshot where supported              |
| Raw Office.js                         | Code and explanation shown verbatim, explicit confirmation every time, time/output limits       |
| Browser `bash`                        | No native process, PC filesystem, direct sockets, dynamic package install, or credential access |

The model may never directly call a confirmation endpoint. A proposal is execution data owned by the UI session; repeated, stale, disconnected, cancelled, logged-out, or wrong-host confirmations fail closed.

## Shared browser runtime

- An in-memory VFS is the default. Optional IndexedDB persistence may store user uploads and skill files only; never credentials or bridge capabilities.
- Paths are normalized below `/home/user` or `/home/skills`; traversal, absolute escape, NUL, oversized names, symlink semantics, and device paths are rejected.
- Per-file, total-byte, file-count, command-output, command-time, and concurrent-command limits are mandatory.
- `SKILL.md` packages are parsed as bounded UTF-8 with strict frontmatter and mounted read-only below `/home/skills/<name>`.
- The skill prompt lists metadata and paths; it does not blindly concatenate unbounded files into every model request.
- Web commands call a fixed WisWork endpoint through the PC bridge or another explicit authenticated bridge route. Arbitrary direct task-pane network access is out of scope.
- File conversions use browser-safe libraries/workers with fixed inputs and output quotas. A command returns `command_unsupported` until its real converter is present; it must not fake success.

## Office.js execution boundary

- Each host adapter checks the active Office host and required API set before calling `Word.run`, `Excel.run`, or `PowerPoint.run`.
- Inputs use exact-object schemas, bounded arrays/strings/ranges, and reject unknown properties.
- Read results are normalized to JSON-safe bounded data before entering Agent history.
- Raw Office.js runs in a hardened JavaScript compartment with an explicit endowment set. `eval`, `Function`, DOM globals, fetch/XHR/WebSocket, storage, cookies, Office auth, and PC bridge objects are unavailable.
- A timeout aborts the agent-side wait; host operations must check cancellation before each `context.sync()` and may not apply a proposal after cancellation.
- Errors returned to the model are stable stage/status codes. Office and upstream exception bodies are not forwarded verbatim.

## Host behavior

### Word

Reads return paragraph indices, styles, list metadata, headings, tables, content controls, sections, and bounded OOXML mappings. Mutations performed through raw Office.js remain confirmation-only. Screenshot support must be feature-detected and otherwise fail closed.

### Excel

Range reads include values, formulas, number formats, and addresses. Mutation proposals summarize affected sheets/ranges and cell counts. Structure deletion and object mutation require elevated confirmation. After confirmation, affected ranges or object inventories are re-read for verification.

### PowerPoint

Reads return slide/shape IDs and geometry. Mutations summarize affected slide IDs and shapes. XML/master edits capture a bounded pre-write snapshot. Successful mutations are followed by `verify_slides` and, where supported, a screenshot check before the Agent may claim completion.

## Failure handling

- Unsupported host/API: `office_api_unsupported`.
- Malformed or oversized tool input: `invalid_tool_input`.
- VFS quota/path violation: `vfs_limit` or `vfs_path_denied`.
- Sandbox policy violation: `sandbox_denied`.
- Command timeout/cancel: `command_timeout` or `cancelled`.
- Stale proposal or changed document selection/state: `proposal_stale`.
- Office exception: stable `office_read_failed`, `office_write_failed`, or `office_verify_failed`.
- PC logout/offline cancels the run, clears proposals and history, and leaves VFS credentials absent.

## Verification

- Contract tests assert the exact tool inventory for each host and absence of other-host tools.
- Unit tests cover schema bounds, cancellation, safe errors, VFS traversal/quotas, sandbox global denial, skill parsing, and proposal immutability/staleness.
- Adapter tests use mocked Office.js request contexts for each read/write path.
- Integration tests run AgentLoop tool calls through each composed host skill and confirmation boundary.
- Build inspection verifies no PC token, OAuth callback, wildcard CSP, native shell dependency, or source map enters the task pane.
- Manual Word/Excel/PowerPoint acceptance covers reads, screenshots, one structured edit, one raw Office.js edit, verification, cancellation, logout, and unsupported API behavior on Windows and macOS.

## Rollback

Host skills and shared browser skills are feature-gated independently. Rollback disables the new registries and restores the existing selection-only skill without changing PC Bridge, authentication, stored user data, or manifests. All proposal and VFS runtime state is disposable.
