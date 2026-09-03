function previewMetadata({ pr, commit, version, builtAt = new Date().toISOString() }) {
  if (!/^[1-9][0-9]{0,5}$/.test(pr ?? '')) throw new Error('Invalid preview PR number')
  if (!/^[a-f0-9]{40}$/.test(commit ?? '')) throw new Error('Invalid preview commit')
  if (!/^\d+\.\d+\.\d+$/.test(version ?? '')) throw new Error('Invalid base version')
  if (typeof builtAt !== 'string' || !Number.isFinite(Date.parse(builtAt))) {
    throw new Error('Invalid preview build time')
  }
  return {
    productName: `WisWork Preview PR${pr}`,
    appId: `com.atominnolab.wiswork.preview.pr${pr}`,
    version: `${version}-pr${pr}.g${commit.slice(0, 7)}`,
    iteration: { mode: 'preview', commit, builtAt, pr: Number(pr) },
  }
}

function createPreviewConfig(base, inputs) {
  const metadata = previewMetadata(inputs)
  return {
    ...base,
    appId: metadata.appId,
    productName: metadata.productName,
    directories: { ...base.directories, output: 'release-preview' },
    publish: null,
    protocols: [],
    fileAssociations: [],
    extraMetadata: {
      ...base.extraMetadata,
      name: `wiswork-preview-pr${inputs.pr}`,
      productName: metadata.productName,
      version: metadata.version,
      wisworkIteration: metadata.iteration,
    },
    mac: {
      ...base.mac,
      target: [{ target: 'dir', arch: ['arm64'] }],
      identity: null,
      notarize: false,
      artifactName: 'WisWork-Preview-${version}-${arch}.${ext}',
    },
    dmg: { ...base.dmg, sign: false },
    afterAllArtifactBuild: undefined,
  }
}

module.exports = { previewMetadata, createPreviewConfig }
