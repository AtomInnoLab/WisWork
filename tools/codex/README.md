# Enhanced mode component release inputs

This directory is internal release infrastructure for the optional Enhanced mode component. The
base WisWork installer does not contain the component, and Standard mode does not require it.

`manifest.json` pins `darwin-arm64`, `darwin-x64`, and `win32-x64` to the official OpenAI
`rust-v0.147.0` app-server package assets. An install first uses the versioned WisWork CDN mirror,
then the exact official GitHub asset as fallback; both sources must match the same reviewed byte
size and SHA-256 digest. Every archive entry is allowlisted and integrity checked. Only
`bin/codex-app-server` (or `.exe`) is retained. Code mode, search, shell, command-runner, and sandbox
helpers are discarded because Enhanced mode exposes none of their generic authorities.

The release workflow must verify the retained executable against the pinned macOS Team Identifier
with strict `codesign` validation or the exact Windows Authenticode publisher, exercise the real
app-server contract on every platform, and pass the package-exclusion and legal-notice gates. The
official 0.147.0 macOS command-line assets are not standalone Gatekeeper-notarized app bundles, so
`spctl` assessment is reserved for the signed/notarized WisWork application artifacts. A checksum
match alone is not release authorization.

The expected team identifier and exact Windows signer-certificate thumbprint/publisher are
release-gate inputs. The Windows identity was calibrated from the reviewed, fixed-digest official
asset and is revalidated by a platform-native Authenticode check. The macOS and Windows release jobs
read the real vendor signature and fail closed on any mismatch. Changing any identity value requires
a reviewed manifest change and must never be accepted dynamically.
