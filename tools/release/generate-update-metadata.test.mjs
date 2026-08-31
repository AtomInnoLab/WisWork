import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import { load } from 'js-yaml'
import { generateUpdateMetadata, parseArgs } from './generate-update-metadata.mjs'

test('generates updater metadata from the final signed artifact bytes', async (context) => {
  const root = await mkdtemp(join(tmpdir(), 'update-metadata-'))
  context.after(() => rm(root, { recursive: true, force: true }))
  const artifact = join(root, 'WisWork-0.6.3-arm64.zip')
  const output = join(root, 'latest-mac.yml')
  const contents = 'final signed zip bytes'
  await writeFile(artifact, contents)

  await generateUpdateMetadata({ artifact, output, version: '0.6.3' })

  const metadata = load(await readFile(output, 'utf8'))
  const sha512 = createHash('sha512').update(contents).digest('base64')
  assert.equal(metadata.version, '0.6.3')
  assert.deepEqual(metadata.files, [
    { url: 'WisWork-0.6.3-arm64.zip', sha512, size: Buffer.byteLength(contents) },
  ])
  assert.equal(metadata.path, 'WisWork-0.6.3-arm64.zip')
  assert.equal(metadata.sha512, sha512)
})

test('rejects unsupported artifacts and incomplete arguments', async (context) => {
  const root = await mkdtemp(join(tmpdir(), 'update-metadata-'))
  context.after(() => rm(root, { recursive: true, force: true }))
  const artifact = join(root, 'WisWork.dmg')
  await writeFile(artifact, 'dmg')

  await assert.rejects(
    generateUpdateMetadata({ artifact, output: join(root, 'latest.yml'), version: '0.6.3' }),
    /unsupported/i,
  )
  assert.throws(() => parseArgs(['--version', '0.6.3']), /expected/i)
})
