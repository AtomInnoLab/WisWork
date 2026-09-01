const GLOBAL_EXTRA_RESOURCES = [
  { from: 'build/THIRD-PARTY-NOTICES.txt', to: 'THIRD-PARTY-NOTICES.txt' },
  { from: '../../node_modules/electron/dist/LICENSES.chromium.html', to: 'LICENSES.chromium.html' },
  { from: '../docs/out', to: 'modules/docs' },
  { from: '../sheets/out', to: 'modules/sheets' },
  { from: '../slides/out', to: 'modules/slides' },
  { from: '../pdf/out', to: 'modules/pdf' },
  { from: '../markdown/out', to: 'modules/markdown' },
  { from: '../../node_modules/@embedpdf/pdfium/dist/pdfium.wasm', to: 'wasm/pdfium.wasm' },
  { from: '../pdf/node_modules/harfbuzzjs/hb-subset.wasm', to: 'wasm/hb-subset.wasm' },
  { from: '../latex/out', to: 'modules/latex' },
  { from: '../../tools/tectonic/manifest.json', to: 'native/tectonic-manifest.json' },
]

function createPackageInputs({ tectonicSource } = {}) {
  const native = (windows) => [
    {
      from: windows
        ? '../sheets/native/xlsx-engine/target/release/xlsx-sidecar.exe'
        : '../sheets/native/xlsx-engine/target/release/xlsx-sidecar',
      to: windows ? 'native/xlsx-sidecar.exe' : 'native/xlsx-sidecar',
    },
    {
      from:
        tectonicSource ?? (windows ? '../latex/native/tectonic.exe' : '../latex/native/tectonic'),
      to: windows ? 'native/tectonic.exe' : 'native/tectonic',
    },
  ]
  return {
    files: ['out/**'],
    extraResources: GLOBAL_EXTRA_RESOURCES.map((entry) => ({ ...entry })),
    macExtraResources: native(false),
    winExtraResources: native(true),
    linuxExtraResources: native(false),
  }
}

module.exports = { createPackageInputs }
