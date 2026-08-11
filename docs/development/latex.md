# LaTeX development and packaging

WisWork packages the LaTeX renderer at
`WisWork.app/Contents/Resources/modules/latex` and the pinned Tectonic 0.16.9 executable at
`WisWork.app/Contents/Resources/native/tectonic`. The 2.8 GB `tectonic-default-bundle-v33` TeX
bundle is deliberately not included in the DMG or ZIP.

## Online indexed bundle and package cache

The compiler manifest in `tools/tectonic/manifest.json` pins the bundle URL, exact byte length,
SHA-256 digest, and license source. Tectonic 0.16.9's legacy indexed `.tar` format is an HTTP range
bundle: it is valid at its official HTTPS URL but is not a supported local bundle file. WisWork
therefore passes only the validated, versioned official URL to Tectonic and does not download or
misreport the complete 2.8 GB tar as a local bundle.

On first compile, Tectonic downloads the index and only the TeX package data required by the
document. Downloaded package data is retained under the controlled user-data path
`latex/tectonic-cache`; on macOS the full default location is
`~/Library/Application Support/WisWork/latex/tectonic-cache`. Later compiles reuse cached package
data but may still access the pinned bundle URL when a document needs uncached files. Packaging CI
downloads only the small Tectonic executable and never downloads or embeds the 2.8 GB bundle.

The renderer and project cannot select an executable, bundle URL, cache directory, environment,
headers, or command-line arguments. Remote bundle mode accepts only the exact pinned
`https://relay.fullyjustified.net/default_bundle_v33.tar` URL. Local `.ttb` and `.zip` bundles
remain supported by the compiler runner in `--only-cached` mode; local `.tar` files are rejected
before process spawn.

## Local sidecar setup

Populate the ignored sidecar cache and publish a verified executable to an explicit path:

```bash
node tools/fetch-tectonic.mjs --platform darwin-arm64 \
  --output "$PWD/apps/latex/native/tectonic"
file apps/latex/native/tectonic
apps/latex/native/tectonic --version
```

The fetcher accepts no caller-provided URL. It validates the manifest-pinned size and SHA-256,
restricts redirect hosts, rejects unsafe archive layouts, and requires exact version output before
publishing the executable. Its archive/extraction cache is `tools/tectonic/.cache/` and is ignored
by Git.

## Cache cleanup

Quit WisWork before cleanup. The following macOS user-data entries are independently removable:

- `latex/compile-cache/`: committed PDF/SyncTeX generations; safe to remove.
- `latex/compile-temp/`: abandoned compile staging; safe to remove while the app is stopped.
- `latex/project-state/`: AI proposal and undo snapshots; removing it discards recovery history.
- `latex/tectonic-cache/`: package data fetched from the pinned indexed bundle; removing it makes
  later compiles fetch required package data again.
- `latex/bundles/`: legacy or future local-compatible bundle assets. The current indexed `.tar`
  runtime does not use the previously downloaded 2.8 GB file, so it may be removed while WisWork
  is stopped.

Never delete a user's LaTeX project while clearing WisWork caches.

## Troubleshooting and current limits

- `file WisWork.app/Contents/Resources/native/tectonic` must report `arm64` for the macOS arm64
  artifact, and `tectonic --version` must report `tectonic 0.16.9`.
- A missing `modules/latex/renderer/index.html` indicates an incomplete application build.
- Network errors require access to the manifest-pinned official bundle host. Do not add alternate
  download URLs or project-controlled mirrors.
- The current editor does not provide arbitrary TeX package installation, shell escape,
  project-supplied executables, or arbitrary network destinations. These are intentional current
  limitations and security boundaries.
- Model credentials and OAuth tokens are unrelated to Tectonic. Never place keys, fixed codes,
  authorization headers, or tokens in this document, the bundle cache, project files, or CI logs.
