import assert from 'node:assert/strict'
import { createRequire } from 'node:module'
import test from 'node:test'
import { readFileSync } from 'node:fs'
import * as yaml from 'js-yaml'

const require = createRequire(import.meta.url)

test('preview identity is isolated and provenance is validated', () => {
  const { previewMetadata } = require('./preview-metadata.cjs')
  const metadata = previewMetadata({
    pr: '123',
    commit: 'a'.repeat(40),
    version: '0.6.14',
    builtAt: '2026-09-03T00:00:00.000Z',
  })
  assert.equal(metadata.productName, 'WisWork Preview PR123')
  assert.equal(metadata.appId, 'com.atominnolab.wiswork.preview.pr123')
  assert.equal(metadata.version, '0.6.14-pr123.aaaaaaa')
  assert.equal(metadata.iteration.mode, 'preview')
  for (const pr of ['0', '-1', '123/../../', '01', '${{ secrets.TOKEN }}']) {
    assert.throws(() => previewMetadata({ pr, commit: 'a'.repeat(40), version: '0.6.14' }))
  }
  assert.throws(() => previewMetadata({ pr: '1', commit: 'main', version: '0.6.14' }))
})

test('automatic preview builds have no signing secrets or release permissions', () => {
  const source = readFileSync(new URL('../../.github/workflows/desktop-preview.yml', import.meta.url), 'utf8')
  const workflow = yaml.load(source)
  assert.ok(workflow.on.pull_request)
  assert.equal(workflow.on.pull_request_target, undefined)
  assert.deepEqual(workflow.permissions, { contents: 'read' })
  assert.ok(!source.includes('secrets.'))
  const steps = workflow.jobs.preview.steps
  const checkout = steps.find((step) => step.uses?.startsWith('actions/checkout@'))
  assert.equal(checkout.with['persist-credentials'], false)
  const upload = steps.find((step) => step.uses?.startsWith('actions/upload-artifact@'))
  assert.equal(upload.with['retention-days'], 7)
  assert.ok(steps.some((step) => step.run?.includes('--mode post-package')))
  assert.ok(steps.some((step) => step.run?.includes('electron-builder.preview.cjs')))
})

test('preview packaging cannot inherit a production publisher or file handlers', () => {
  const { createPreviewConfig } = require('./preview-metadata.cjs')
  const config = createPreviewConfig(
    {
      publish: [{ provider: 'github' }],
      protocols: [{ schemes: ['wiswork'] }],
      fileAssociations: [{ ext: 'pptx' }],
      mac: { notarize: true, extraResources: [] },
      afterAllArtifactBuild: 'sign.js',
    },
    { pr: '12', commit: 'b'.repeat(40), version: '0.6.14', builtAt: '2026-09-03T00:00:00.000Z' },
  )
  assert.equal(config.publish, null)
  assert.deepEqual(config.protocols, [])
  assert.deepEqual(config.fileAssociations, [])
  assert.equal(config.mac.notarize, false)
  assert.equal(config.mac.identity, null)
  assert.equal(config.afterAllArtifactBuild, undefined)
  assert.equal(config.extraMetadata.wisworkIteration.mode, 'preview')
  assert.equal(config.extraMetadata.productName, config.productName)
})
