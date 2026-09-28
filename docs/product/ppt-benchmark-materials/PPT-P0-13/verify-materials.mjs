import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { PNG } from 'pngjs'

const root = dirname(fileURLToPath(import.meta.url))
const lines = (await readFile(join(root, 'SHA256SUMS'), 'utf8')).trim().split('\n')
assert.equal(lines.length, 15)
const hashes = new Map()
for (const line of lines) {
  const match = /^([a-f0-9]{64})  ([A-Za-z0-9./-]+)$/.exec(line)
  assert.ok(match, 'invalid SHA256SUMS row')
  const [, expected, name] = match
  assert.ok(!name.includes('..') && !hashes.has(name), 'invalid or repeated file name')
  const actual = createHash('sha256')
    .update(await readFile(join(root, name)))
    .digest('hex')
  assert.equal(actual, expected, `changed ${name}`)
  hashes.set(name, actual)
}
const rights = JSON.parse(await readFile(join(root, 'asset-rights.json'), 'utf8'))
assert.equal(rights.version, 1)
assert.equal(rights.assets.length, 12)
const imageDigests = new Set()
for (const [index, asset] of rights.assets.entries()) {
  const name = `images/schematic-${String(index + 1).padStart(2, '0')}.png`
  assert.equal(asset.file, name)
  assert.equal(asset.sha256, hashes.get(name))
  assert.ok(asset.usePermission && asset.attribution)
  imageDigests.add(asset.sha256)
  const image = PNG.sync.read(await readFile(join(root, name)))
  assert.deepEqual([image.width, image.height], [960, 540])
}
assert.equal(imageDigests.size, 12)
const source = await readFile(join(root, 'deardorff-2020-article.pdf'))
const original = await readFile(join(root, '../PPT-P0-01/deardorff-2020-article.pdf'))
assert.ok(source.equals(original), 'source PDF differs from original candidate')
const zip = await JSZip.loadAsync(
  await readFile(join(root, 'wiswork-image-dense-research-draft.pptx')),
)
const slides = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
const charts = Object.keys(zip.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))
assert.equal(slides.length, 8)
assert.equal(charts.length, 2)
const pictureCounts = []
for (let page = 1; page <= 8; page++) {
  const xml = await zip.file(`ppt/slides/slide${page}.xml`)?.async('string')
  assert.ok(xml?.includes('<a:t>'), `blank slide ${page}`)
  pictureCounts.push((xml.match(/<p:pic>/g) ?? []).length)
}
assert.deepEqual(pictureCounts, [2, 2, 2, 2, 1, 1, 2, 2])
process.stdout.write(
  'PPT-P0-13 candidate verified: 8 pages, 12 licensed images, 14 placements, 2 native charts.\n',
)
