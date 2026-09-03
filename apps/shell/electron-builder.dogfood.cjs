// Remove release environment before loading its configuration (including hooks).
delete process.env.WISWORK_UPDATE_PROVIDER
delete process.env.WISWORK_UPDATE_URL
delete process.env.WISWORK_MAC_X64
delete process.env.WISWORK_TECTONIC_SOURCE
process.env.WISWORK_UNSIGNED_MAC_BUILD = '1'
const base = require('./electron-builder.cjs')
const { existsSync } = require('node:fs')
const { join } = require('node:path')
const { execFileSync } = require('node:child_process')
const commit = process.env.WISWORK_ITERATION_COMMIT
const builtAt = process.env.WISWORK_ITERATION_BUILT_AT
if (!/^[a-f0-9]{7,40}$/.test(commit ?? '') || !Number.isFinite(Date.parse(builtAt ?? ''))) {
  throw new Error('Dogfood packaging requires commit and build time metadata')
}
module.exports = {
  ...base,
  appId: 'com.atominnolab.wiswork.dogfood',
  productName: 'WisWork Dogfood',
  directories: { output: 'release-dogfood' },
  protocols: [],
  fileAssociations: [],
  publish: null,
  extraMetadata: {
    name: 'wiswork-dogfood',
    productName: 'WisWork Dogfood',
    wisworkIteration: { mode: 'dogfood', commit, builtAt },
  },
  mac: {
    ...base.mac,
    target: [{ target: 'dir', arch: [process.arch] }],
    identity: null,
    notarize: false,
  },
  afterAllArtifactBuild: undefined,
  beforePack: async (context) => {
    await base.beforePack(context)
    for (const rel of [
      '../latex/native/tectonic',
      '../sheets/native/xlsx-engine/target/release/xlsx-sidecar',
    ]) {
      const path = join(__dirname, rel)
      if (!existsSync(path)) throw new Error(`Missing required native asset: ${rel}`)
      const archs = execFileSync('lipo', ['-archs', path], { encoding: 'utf8' }).trim().split(/\s+/)
      if (!archs.includes(process.arch === 'x64' ? 'x86_64' : 'arm64'))
        throw new Error(`Native asset has wrong architecture: ${rel}`)
    }
  },
}
