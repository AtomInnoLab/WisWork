import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'

const root = dirname(fileURLToPath(import.meta.url))
const digest = (value) => createHash('sha256').update(value).digest('hex')
const entries = (await readFile(join(root, 'SHA256SUMS'), 'utf8')).trim().split('\n')
assert.equal(entries.length, 16)
const hashes = new Map()
for (const entry of entries) {
  const match = /^([a-f0-9]{64})  ([A-Za-z0-9./-]+)$/.exec(entry)
  assert.ok(match, 'invalid SHA256SUMS entry')
  const [, expected, path] = match
  assert.ok(!path.includes('..') && !hashes.has(path), 'unsafe or duplicate file path')
  assert.equal(digest(await readFile(join(root, path))), expected, `changed ${path}`)
  hashes.set(path, expected)
}
const previous = (await readFile(join(root, '../PPT-P0-13/SHA256SUMS'), 'utf8')).trim().split('\n')
for (const entry of previous) {
  const [, hash, path] = /^([a-f0-9]{64})  ([A-Za-z0-9./-]+)$/.exec(entry) ?? []
  assert.equal(hashes.get(path), hash, `source version changed: ${path}`)
}
const plan = JSON.parse(await readFile(join(root, 'page-plan.json'), 'utf8'))
assert.equal(plan.version, 1)
assert.equal(plan.status, 'candidate_pending_scientific_and_layout_review')
assert.equal(plan.source.file, 'deardorff-2020-article.pdf')
assert.equal(plan.pages.length, 8)
assert.deepEqual(
  plan.pages.map((page) => page.id),
  ['p01', 'p02', 'p03', 'p04', 'p05', 'p06', 'p07', 'p08'],
)
assert.deepEqual(
  plan.pages.filter((page) => page.revisionTarget).map((page) => page.id),
  ['p04'],
)
assert.deepEqual(plan.preserveUnchangedPageIds, ['p01', 'p02', 'p03', 'p05', 'p06', 'p07', 'p08'])
assert.deepEqual(plan.recheckAfterRevision, ['p04', 'p08'])
for (const [index, page] of plan.pages.entries()) {
  assert.ok(page.title && page.role)
  assert.ok(
    page.sourcePages?.every((number) => Number.isInteger(number) && number >= 1 && number <= 11),
  )
  assert.ok(
    page.dependsOn?.every((id) => plan.pages.slice(0, index).some((earlier) => earlier.id === id)),
  )
}
const rights = JSON.parse(await readFile(join(root, 'asset-rights.json'), 'utf8'))
assert.equal(rights.assets.length, 12)
for (const [index, asset] of rights.assets.entries()) {
  assert.equal(asset.file, `images/schematic-${String(index + 1).padStart(2, '0')}.png`)
  assert.equal(asset.sha256, hashes.get(asset.file))
}
const zip = await JSZip.loadAsync(
  await readFile(join(root, 'wiswork-image-dense-research-draft.pptx')),
)
assert.equal(
  Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length,
  8,
)
assert.equal(
  Object.keys(zip.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length,
  2,
)
const pictureCounts = []
for (const [index, page] of plan.pages.entries()) {
  const xml = await zip.file(`ppt/slides/slide${index + 1}.xml`)?.async('string')
  assert.ok(xml?.includes(page.title), `slide title differs for ${page.id}`)
  pictureCounts.push((xml.match(/<p:pic>/g) ?? []).length)
}
assert.deepEqual(pictureCounts, [2, 2, 2, 2, 1, 1, 2, 2])
process.stdout.write(
  'PPT-P0-18 candidate verified: 8-page frozen report, 7 preserved pages, 1 revision target.\n',
)
