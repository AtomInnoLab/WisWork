# LaTeX development and packaging

WisWork packages the LaTeX renderer at
`WisWork.app/Contents/Resources/modules/latex` and the pinned Tectonic 0.16.9 executable at
`WisWork.app/Contents/Resources/native/tectonic`. The 2.8 GB `tectonic-default-bundle-v33` TeX
bundle is deliberately not included in the DMG or ZIP.

## First compile and bundle cache

The compiler manifest in `tools/tectonic/manifest.json` pins the bundle URL, exact byte length,
SHA-256 digest, and license source. `BundleInstaller` supports resumable, bounded downloading,
hash verification, an installation lock, and atomic publication. The desktop runtime consumes the
verified bundle at the user-data path `latex/bundles/tectonic-default-bundle-v33.tar`; on macOS the
full default location is
`~/Library/Application Support/WisWork/latex/bundles/tectonic-default-bundle-v33.tar`.

On first compile, the main process checks the cache and starts one shared `BundleInstaller`
download when the bundle is missing. The Compile panel reports bounded progress and Cancel aborts
the download. A verified file is published atomically as
`latex/bundles/tectonic-default-bundle-v33.tar`; later launches validate and reuse that concrete tar
offline without another download. Do not copy an unverified TeX bundle into the cache. Packaging
CI downloads only the small Tectonic executable and never downloads or embeds the 2.8 GB bundle.

After a verified bundle is present, compilation runs Tectonic with `--only-cached` and an explicit
`--bundle` path. Recompilation is offline: it does not let a project select an executable, bundle
URL, environment, headers, or command-line arguments. A document that needs files absent from the
pinned bundle still fails deterministically rather than fetching arbitrary content.

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
- `latex/bundles/`: the verified TeX bundle; removing it disables offline compilation and makes
  the next compile download and verify the pinned bundle again.

Never delete a user's LaTeX project while clearing WisWork caches.

## Troubleshooting and current limits

- `file WisWork.app/Contents/Resources/native/tectonic` must report `arm64` for the macOS arm64
  artifact, and `tectonic --version` must report `tectonic 0.16.9`.
- A missing `modules/latex/renderer/index.html` indicates an incomplete application build.
- A missing or wrong-size bundle is not repaired silently. Check the manifest byte length and
  SHA-256 without logging its contents or adding alternate download URLs.
- The current editor does not provide arbitrary TeX package installation, shell escape,
  project-supplied executables, or network access after the verified bundle is installed. These
  are intentional current limitations and security boundaries.
- Model credentials and OAuth tokens are unrelated to Tectonic. Never place keys, fixed codes,
  authorization headers, or tokens in this document, the bundle cache, project files, or CI logs.
