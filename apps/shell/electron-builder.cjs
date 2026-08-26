/**
 * electron-builder configuration (moved out of package.json "build" so the
 * auto-update feed URL can be injected at build time instead of living in
 * the repo).
 *
 * WISWORK_UPDATE_URL — public base URL of the update channel (the generic
 * provider prefix that serves latest.yml / latest-mac.yml). Required for
 * release builds; CI provides it as a repository secret. For local release
 * builds put it in apps/shell/electron-builder.env (gitignored) — the
 * electron-builder CLI loads that file automatically.
 *
 * When the variable is unset (forks, PR smoke builds, plain local packaging)
 * the publish config is omitted: electron-builder then bakes no
 * app-update.yml into the app and in-app auto-update stays disabled.
 */

const { execFileSync } = require('node:child_process')
const { existsSync } = require('node:fs')
const { join } = require('node:path')

const updateUrl = process.env.WISWORK_UPDATE_URL
// Explicitly scoped to disposable test artifacts. Production release jobs
// must omit this variable so signing and notarization remain fail-closed.
const unsignedMacBuild = process.env.WISWORK_UNSIGNED_MAC_BUILD === '1'
const tectonicSource = process.env.WISWORK_TECTONIC_SOURCE
// WISWORK_MAC_X64=1 — opt into packaging the Intel (x64) dmg/zip alongside
// arm64. Off by default: Intel packages must only ever ship signed with the
// company certificate (planned dual-track pipeline), so the current release
// pipeline stays arm64-only and never produces a personally-signed Intel
// artifact. The downstream layout (feed archive name, WisWork-intel.dmg
// alias) keys off which dmgs exist, so flipping this flag is the single
// switch.
const includeMacX64 = process.env.WISWORK_MAC_X64 === '1'

// LICENSES.chromium.html only exists after the Electron binary download —
// since Electron 42 that no longer happens during `npm ci` (the postinstall
// script was replaced by the lazy `install-electron` bin), and electron-builder
// exits 0 on a missing extraResources source, so without this check the
// installer would silently ship without the Chromium license.
for (const rel of [
  '../../node_modules/electron/dist/LICENSES.chromium.html',
  '../../node_modules/@embedpdf/pdfium/dist/pdfium.wasm',
  '../pdf/node_modules/harfbuzzjs/hb-subset.wasm',
]) {
  if (!existsSync(join(__dirname, rel))) {
    throw new Error(
      `electron-builder extraResources source missing: ${rel} (npm hoisting changed?)`,
    )
  }
}

// The module trees are electron-vite outputs produced by build:all; a missing
// one means that module's build did not run or failed. electron-builder only
// logs "file source doesn't exist" for an absent extraResources source and
// still exits 0, so without this the installer launches normally and is simply
// missing that editor — it surfaces only when a user opens the tab.
//
// Runs from the beforePack hook, not at module load: gen-third-party-notices
// requires this config to read extraResources, and the dist:* scripts run
// notices before build:all, when the out dirs legitimately don't exist yet.
// When the mac build packages BOTH arches (WISWORK_MAC_X64=1) its
// extraResources entry is a single path shared by the two packs, so the
// sidecar there must be a lipo fat binary — a host-arch-only build (the plain
// `native:build` dev path) would silently ship an arm64 sidecar inside the
// Intel dmg, where every workbook open fails. Runs from beforePack, dual-arch
// mac packs only.
function assertUniversalSidecar() {
  const sidecar = join(__dirname, '../sheets/native/xlsx-engine/target/release/xlsx-sidecar')
  if (!existsSync(sidecar)) {
    throw new Error(
      `mac extraResources source missing: ${sidecar} (run "npm run native:build:universal -w @wiswork/sheets" first)`,
    )
  }
  const archs = execFileSync('lipo', ['-archs', sidecar], { encoding: 'utf8' }).trim().split(/\s+/)
  for (const want of ['x86_64', 'arm64']) {
    if (!archs.includes(want)) {
      throw new Error(
        `xlsx-sidecar is [${archs.join(', ')}] but both mac arch packages ship it — ` +
          'run "npm run native:build:universal -w @wiswork/sheets" before packaging mac',
      )
    }
  }
}

function assertModuleTreesPresent() {
  for (const rel of [
    '../docs/out',
    '../sheets/out',
    '../slides/out',
    '../pdf/out',
    '../markdown/out',
    '../latex/out',
  ]) {
    if (!existsSync(join(__dirname, rel))) {
      throw new Error(
        `electron-builder extraResources source missing: ${rel} (run npm run build:all first)`,
      )
    }
  }
}

/** @type {import('electron-builder').Configuration} */
const config = {
  appId: 'com.atominnolab.wiswork',
  productName: 'WisWork',
  artifactName: 'WisWork-${version}-${arch}.${ext}',
  // Resolve from the installed package so runtime dependency updates cannot
  // leave a stale hard-coded packaging pin.
  electronVersion: require('electron/package.json').version,
  // The shell is the authoritative installed owner. Standalone builds using the same scheme conflict if installed together.
  protocols: [{ name: 'WisWork OAuth Callback', schemes: ['wiswork'] }],
  directories: {
    output: 'release',
  },
  files: ['out/**'],
  extraResources: [
    {
      from: 'build/THIRD-PARTY-NOTICES.txt',
      to: 'THIRD-PARTY-NOTICES.txt',
    },
    {
      from: '../../node_modules/electron/dist/LICENSES.chromium.html',
      to: 'LICENSES.chromium.html',
    },
    {
      from: '../docs/out',
      to: 'modules/docs',
    },
    {
      from: '../sheets/out',
      to: 'modules/sheets',
    },
    {
      from: '../slides/out',
      to: 'modules/slides',
    },
    {
      from: '../pdf/out',
      to: 'modules/pdf',
    },
    {
      from: '../markdown/out',
      to: 'modules/markdown',
    },
    // PDF text editing engines: the bundled main resolves these under
    // Resources/wasm when node_modules is absent (apps/pdf/src/main/wasm-path.ts)
    {
      from: '../../node_modules/@embedpdf/pdfium/dist/pdfium.wasm',
      to: 'wasm/pdfium.wasm',
    },
    {
      from: '../pdf/node_modules/harfbuzzjs/hb-subset.wasm',
      to: 'wasm/hb-subset.wasm',
    },
    {
      from: '../latex/out',
      to: 'modules/latex',
    },
    {
      from: '../../tools/tectonic/manifest.json',
      to: 'native/tectonic-manifest.json',
    },
  ],
  // `mimeType` is read only by the Linux target, where it becomes the
  // desktop entry's MimeType= list; associations without it are dropped
  // there. macOS and Windows ignore the field and key off `ext`.
  fileAssociations: [
    {
      ext: 'docx',
      name: 'Word Document',
      role: 'Editor',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    },
    {
      ext: 'xlsx',
      name: 'Excel Workbook',
      role: 'Editor',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
    {
      ext: 'pptx',
      name: 'PowerPoint Presentation',
      role: 'Editor',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    },
    {
      ext: 'xls',
      name: 'Excel 97-2003 Workbook',
      role: 'Editor',
      mimeType: 'application/vnd.ms-excel',
    },
    {
      ext: 'csv',
      name: 'CSV Document',
      role: 'Editor',
      mimeType: 'text/csv',
    },
    {
      ext: 'pdf',
      name: 'PDF Document',
      role: 'Editor',
      mimeType: 'application/pdf',
    },
    {
      ext: 'md',
      name: 'Markdown Document',
      role: 'Editor',
      mimeType: 'text/markdown',
    },
    {
      ext: 'markdown',
      name: 'Markdown Document',
      role: 'Editor',
      mimeType: 'text/markdown',
    },
  ],
  npmRebuild: false,
  mac: {
    // Two separate arch packages (NOT universal): arm64 keeps the exact
    // artifact names and update-feed entries it always had, x64 (opt-in via
    // WISWORK_MAC_X64=1, see includeMacX64 above) adds Intel support with
    // electron-builder's default arch-less names. Both zips land in one latest-mac.yml and
    // electron-updater picks by process.arch. Dual-arch packs ship the same
    // lipo fat xlsx-sidecar (see assertUniversalSidecar above).
    target: [
      { target: 'dmg', arch: includeMacX64 ? ['arm64', 'x64'] : ['arm64'] },
      { target: 'zip', arch: includeMacX64 ? ['arm64', 'x64'] : ['arm64'] },
    ],
    category: 'public.app-category.productivity',
    hardenedRuntime: true,
    gatekeeperAssess: false,
    entitlements: 'build/entitlements.mac.plist',
    entitlementsInherit: 'build/entitlements.mac.plist',
    ...(unsignedMacBuild ? { identity: null, notarize: false } : { notarize: true }),
    extraResources: [
      {
        from: '../sheets/native/xlsx-engine/target/release/xlsx-sidecar',
        to: 'native/xlsx-sidecar',
      },
      {
        from: tectonicSource ?? '../latex/native/tectonic',
        to: 'native/tectonic',
      },
    ],
  },
  win: {
    target: [
      {
        target: 'nsis',
        arch: ['x64'],
      },
    ],
    extraResources: [
      {
        from: '../sheets/native/xlsx-engine/target/x86_64-pc-windows-gnu/release/xlsx-sidecar.exe',
        to: 'native/xlsx-sidecar.exe',
      },
      {
        from: tectonicSource ?? '../latex/native/tectonic.exe',
        to: 'native/tectonic.exe',
      },
    ],
  },
  // Unlike win (which cross-compiles the sidecar to an explicit target
  // triple), linux takes it from cargo's host-native target/release/ — the
  // same source mac uses. So no `arch` is pinned here: electron-builder
  // defaults to the build host's architecture, which is the only one the
  // sidecar was actually built for. Packaging arm64 on an x64 host, or the
  // reverse, needs a matching `cargo build --target` first.
  linux: {
    // AppImage (self-contained, any distro) + deb (apt install, pulls in the
    // GTK/NSS runtime deps) + rpm (dnf/zypper install on Fedora / RHEL /
    // openSUSE). Default artifact names are kept on purpose —
    // WisWork-<v>.AppImage / wiswork_<v>_amd64.deb.
    target: [
      { target: 'AppImage', arch: ['x64'] },
      { target: 'deb', arch: ['x64'] },
      { target: 'rpm', arch: ['x64'] },
    ],
    // deb control metadata; values match the manually published 0.5.149 deb
    // so apt sees the new packages as the same lineage. Homepage comes from
    // package.json "homepage"; the Package field is pinned in the deb block
    // below (packageName is a per-target option, rejected here by the schema).
    maintainer: 'AtomInnoLab <noreply@atominnolab.com>',
    vendor: 'AtomInnoLab',
    category: 'Office',
    icon: 'build/icon.png',
    // mac and win name the binary from productName; linux instead derives it
    // from package.json "name", and the scoped package name does not produce
    // the desired binary name. Setting it explicitly also makes the
    // generated wiswork.desktop match the WM_CLASS Electron reports (it
    // takes that from the executable basename), so the running window links
    // back to its launcher entry.
    executableName: 'wiswork',
    // Electron takes its X11 app_id from package.json "desktopName"
    // (wiswork.desktop); syncDesktopName makes electron-builder name the
    // .desktop file and its StartupWMClass from the same value. Without it
    // StartupWMClass falls back to productName ("WisWork"), which does not
    // match the "wiswork" WM_CLASS the window actually reports — and X11
    // compares case-sensitively, so the taskbar shows an unlinked window.
    syncDesktopName: true,
    extraResources: [
      {
        from: '../sheets/native/xlsx-engine/target/release/xlsx-sidecar',
        to: 'native/xlsx-sidecar',
      },
      {
        from: tectonicSource ?? '../latex/native/tectonic',
        to: 'native/tectonic',
      },
    ],
  },
  // Pin package and artifact names so the npm scope does not leak into Linux
  // package metadata and successive WisWork releases remain upgrade-compatible.
  deb: {
    artifactName: 'wiswork_${version}_${arch}.deb',
    packageName: 'wiswork',
  },
  // Same scoped-package naming problem as deb: spell the artifact name
  // out (${arch} expands to the rpm arch string, x86_64) and pin the rpm
  // Package name so dnf/zypper treat successive releases as upgrades of the
  // same package. Like deb, rpm installs run no in-app updater — users
  // upgrade with `dnf install ./<new>.rpm`. Packaging needs rpmbuild on the
  // build host (the `rpm` apt package on Ubuntu; CI installs it).
  //
  // publish: null (explicit) keeps the rpm out of the electron-updater feed
  // and off the CDN entirely: the rpm is a GitHub-Release download only, so
  // latest-linux.yml keeps listing exactly what the CDN pipeline uploads
  // (AppImage + deb) and the promote workflow needs no rpm alias.
  rpm: {
    artifactName: 'wiswork-${version}.${arch}.rpm',
    packageName: 'wiswork',
    publish: null,
  },
  nsis: {
    oneClick: false,
    allowToChangeInstallationDirectory: true,
  },
  beforePack: async (context) => {
    assertModuleTreesPresent()
    if (context.electronPlatformName === 'darwin' && includeMacX64) assertUniversalSidecar()
  },
  dmg: {
    sign: !unsignedMacBuild,
  },
  afterAllArtifactBuild: unsignedMacBuild ? undefined : 'build/notarize-dmg.js',
}

if (updateUrl) {
  config.publish = [
    {
      provider: 'generic',
      url: updateUrl.replace(/\/+$/, ''),
      channel: 'latest',
    },
  ]
}

module.exports = config
