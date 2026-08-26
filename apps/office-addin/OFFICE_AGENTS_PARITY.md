# office-agents capability parity

Baseline: the `main` branch READMEs in `hewliyang/office-agents`, reviewed on 2026-08-23.
“Supported” means the production skill advertises the operation and has a bounded Office.js or
worker-backed implementation. It does not mean WisWork executes arbitrary JavaScript.

## Host tools

| Host       | Reference tools                                                                                                                                                                                                                      | WisWork status                                                                                                                                                                                               |
| ---------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Word       | `get_document_text`, `get_document_structure`, `get_ooxml`, `screenshot_document`                                                                                                                                                    | Supported                                                                                                                                                                                                    |
| Word       | `execute_office_js`                                                                                                                                                                                                                  | Safe subset: strict declarative `insert_text` and `replace_all`; raw JavaScript is intentionally rejected                                                                                                    |
| Word       | —                                                                                                                                                                                                                                    | WisWork additionally provides `write_document`, which atomically converts bounded Markdown to native Word headings, paragraphs, tables, and paragraph inline formatting through Office.js; lists fail closed |
| Excel      | `get_cell_ranges`, `get_range_as_csv`, `search_data`, `screenshot_range`, `get_all_objects`, `set_cell_range`, `clear_cell_range`, `copy_to`, `modify_sheet_structure`, `modify_workbook_structure`, `resize_range`, `modify_object` | Supported, subject to the documented Office API-set gates                                                                                                                                                    |
| Excel      | `eval_officejs`                                                                                                                                                                                                                      | Safe subset: strict declarative operations only; raw JavaScript is intentionally rejected                                                                                                                    |
| PowerPoint | `screenshot_slide`, `list_slide_shapes`, `read_slide_text`, `verify_slides`, `edit_slide_text`, `edit_slide_xml`, `edit_slide_chart`, `edit_slide_master`, `duplicate_slide`                                                         | Supported, subject to the documented Office API-set gates                                                                                                                                                    |
| PowerPoint | `execute_office_js`                                                                                                                                                                                                                  | Safe subset: strict declarative text/geometry/shape operations only; raw JavaScript is intentionally rejected                                                                                                |

## Shared and custom commands

| Reference surface                                             | WisWork status                                                                                                                                                                                     |
| ------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `read`, `bash`                                                | Supported by the session VFS and bounded command registry; this is not a native shell                                                                                                              |
| `pdf-to-text`, `pdf-to-images`, `docx-to-text`, `xlsx-to-csv` | Supported in a terminateable bounded conversion worker                                                                                                                                             |
| Excel `csv-to-sheet`, `sheet-to-csv`, `image-to-sheet`        | Supported when ExcelApi 1.9 is available                                                                                                                                                           |
| PowerPoint `insert-image`                                     | Supported when PowerPointApi 1.8 is available                                                                                                                                                      |
| `web-search`, `web-fetch`, Word `image-search`                | Not shipped. The task pane contains a capability-scoped client, but production does not advertise it until a fixed retrieval endpoint and reviewed SSRF/DNS-rebinding attestation are compiled in. |

## Deliberate semantic differences

- WisWork never runs model-provided JavaScript, `eval`, `Function`, native shell commands, or ambient
  Office authority. The reference raw Office.js escape hatches are therefore name-compatible safe
  subsets, not semantic parity.
- All WisWork mutations require a visible confirmation, stale-state check, bounded execution, and
  semantic verification. Unsupported API sets fail closed instead of simulating success.
- `write_document` is a WisWork extension. Markdown links, images, raw HTML, scripts, and unknown
  constructs remain literal text; the allowlisted subset is headings, paragraphs, flat ordered or
  plain-text pipe tables, paragraph bold, italic, and inline code. Ordered/unordered lists and
  formatted table cells fail closed rather than relying on non-transactional numbering operations
  or silently losing formatting.
