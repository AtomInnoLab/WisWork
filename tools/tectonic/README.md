# Tectonic assets

`manifest.json` is the sole source of network locations used to package the Tectonic sidecar and
install its TeX bundle. Asset URLs, byte lengths, SHA-256 digests, versions, and license sources
must be reviewed and committed together. Never use a `latest` URL or add credentials/query tokens.

The checked-in manifest is generated only after upstream bytes have been independently verified.
Until every required value is known, do not add placeholder assets to the manifest.

To populate the local, ignored cache after the manifest is committed:

```sh
node tools/fetch-tectonic.mjs --platform darwin-arm64 --output /explicit/path/to/tectonic
```

The fetcher accepts no URL argument. It permits an initial request only to `github.com` and a
manual redirect only to `release-assets.githubusercontent.com`, enforces a five-minute timeout and
the pinned maximum byte length, checks SHA-256, fsyncs a unique temporary file, and atomically
renames it. It then validates the archive contains only the expected executable, always rebuilds
from the verified archive into a staging directory, and requires exact `tectonic --version` output
before a backup-and-swap publish. Stable JSON logs contain only an asset ID, byte count, or error
code; URLs and upstream error text are intentionally omitted.

The optional `--output` path receives only the verified executable and is intended for packaging.
The cache contains downloaded upstream archives and verified executables and must not be committed.
