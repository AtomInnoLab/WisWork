# Enhanced mode component release inputs

This directory is internal release infrastructure for the optional Enhanced mode component. The
base WisWork installer does not contain the component, and Standard mode does not require it.

`manifest.json` pins the one supported pilot target (`darwin-arm64`) to the official OpenAI release
tag `rust-v0.147.0`. The archive is downloaded only after an explicit install request. Every archive
entry is allowlisted and integrity checked; only `bin/codex` and its sibling
`bin/codex-code-mode-host` are retained. The search and shell resources are verified during archive
inspection and then discarded because the pilot disables shell and arbitrary filesystem tools.

The release workflow must verify the vendor signature of both retained Mach-O files, exercise the
real app-server and document-tool contracts, and pass the legal notice gate. A checksum match alone
is not release authorization.
