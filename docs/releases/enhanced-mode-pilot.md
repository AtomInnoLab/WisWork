# Enhanced mode pilot and rollback

WisWork continues to use **Standard mode** by default. **Enhanced mode** is an optional macOS arm64
pilot for LaTeX projects. Downloading it does not enable it, and enabling it takes effect only after
WisWork restarts. The base installer, Standard mode, and every existing editor remain usable when
the optional component is missing, unsupported, cancelled, or fails verification.

## Install and enable

1. Open the account menu and choose **Download** on the **Enhanced mode** row.
2. WisWork downloads the pinned optional component into its private application-data cache.
3. WisWork verifies the archive and every retained file before reporting it ready.
4. Choose **Enable after restart**, then restart WisWork.

The pilot supports only Apple silicon Macs (`darwin-arm64`). It does not use a system installation,
search `PATH`, or accept a custom executable path in a release build. A missing or modified
component produces an install-required message before a turn begins. WisWork never changes modes in
the middle of a turn.

## Roll back or remove

Choose **Switch to Standard mode** and restart WisWork. This flag-only rollback does not change or
migrate documents, saved chat history, proposals, snapshots, or compiler state. The optional
component can be removed after the restarted process is running in Standard mode.

If Enhanced mode stops during a turn, WisWork reports the failure visibly and does not silently
repeat the prompt or tool action in Standard mode. Start a new turn after explicitly switching
modes.

## Privacy and permissions

Enhanced mode uses the existing WisPaper login and WisUsage model service. Renderers receive no
model credential. Document changes still use the existing proposal confirmation, revision guard,
snapshot, validation, compile, and undo path. The optional runtime receives no general shell,
filesystem, document-directory, or network capability.

## Release gates

An Enhanced mode release is blocked unless the macOS arm64 workflow passes all of the following:

- pinned archive and extracted-file SHA-256/size checks;
- strict vendor code-signing and Gatekeeper assessment for both retained executables;
- exact version probe and real local app-server/document-tool contract tests;
- protocol schema, security matrix, lifecycle suite, and LaTeX parity threshold;
- Apache-2.0, upstream NOTICE, V8, and runtime-host license notices.

The base WisWork package intentionally contains none of the optional native executables.

## Internal implementation and third-party attribution

Enhanced mode currently uses OpenAI Codex CLI 0.147.0 (`codex app-server`) downloaded from the
official `rust-v0.147.0` GitHub release. Codex is licensed under Apache-2.0. Its legally required
name, source, NOTICE, and dependency licenses appear in WisWork's third-party notices; this internal
implementation name is not a product setting or user-facing mode name.
