# PPT preflight without packaging

Run the complete functional gate before creating a desktop installer:

```bash
WISWORK_REAL_WISUSAGE_TOKEN='<temporary test token>' npm run test:ppt-preflight
```

On macOS arm64/x64, the runner automatically finds the verified Enhanced component installed by
WisWork under the current user's Application Support directory. To use another reviewed binary,
set `WISWORK_CODEX_INTEGRATION_EXECUTABLE` to its absolute path.

The gate fails unless every stage succeeds:

1. Renderer → Shell → Codex Bridge production tool-catalog contracts.
2. Standard mode with live WisUsage, producing a complete multi-page PPTX.
3. Enhanced mode with the real pinned Codex process, live WisUsage protocol, user-confirmed
   mutation flow, canonical slide transactions, readback, and PPTX export.
4. Built-from-source Electron Slides acceptance rendering. No DMG, ZIP, or installer is created.

The production contract also rejects the old repeated title/body-only builder. A designed deck can
carry a dark or light theme, five bounded layout families (`cover`, `split_image`, `cards`,
`timeline`, and `statement`), and HTTPS imagery selected through `image_search`. The focused gate
checks theme colors, layout diversity, image placement, and export style preservation before any
live model run.

Artifacts and a credential-free report are written to `test-results/ppt-preflight/`:

- `standard-llm.pptx`
- `enhanced-onboarding.pptx`
- `report.json`

The report records only stage names, exit codes, durations, artifact sizes, and SHA-256 digests. It
does not include prompts, document content, tokens, component paths, or credentials. A failed stage
stops the gate immediately and writes a failure report; skipped live tests are never treated as a
pass.
