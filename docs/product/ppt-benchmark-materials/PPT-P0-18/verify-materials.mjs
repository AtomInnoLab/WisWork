import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { pdfToPages } from '../../../../packages/file-parse/src/pdf.ts'
import { checkPresentationChartData } from '@wiswork/pptx-engine/presentation-chart-data'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'

const root = dirname(fileURLToPath(import.meta.url))
const digest = (value) => createHash('sha256').update(value).digest('hex')
const entries = (await readFile(join(root, 'SHA256SUMS'), 'utf8')).trim().split('\n')
assert.equal(entries.length, 30)
const hashes = new Map()
for (const entry of entries) {
  const match = /^([a-f0-9]{64}) {2}([A-Za-z0-9./-]+)$/.exec(entry)
  assert.ok(match, 'invalid SHA256SUMS entry')
  const [, expected, path] = match
  assert.ok(!path.includes('..') && !hashes.has(path), 'unsafe or duplicate file path')
  assert.equal(digest(await readFile(join(root, path))), expected, `changed ${path}`)
  hashes.set(path, expected)
}
const previous = (await readFile(join(root, '../PPT-P0-13/SHA256SUMS'), 'utf8')).trim().split('\n')
for (const entry of previous) {
  const [, hash, path] = /^([a-f0-9]{64}) {2}([A-Za-z0-9./-]+)$/.exec(entry) ?? []
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
const productionPlan = parsePresentationPlan(
  JSON.parse(await readFile(join(root, 'reference-plan.json'), 'utf8')),
)
const parentDeck = JSON.parse(await readFile(join(root, 'reference-deck.json'), 'utf8'))
const revisedDeck = JSON.parse(await readFile(join(root, 'revised-page-deck.json'), 'utf8'))
const manifest = JSON.parse(await readFile(join(root, 'page-artifact-manifest.json'), 'utf8'))
assertDeckMatchesPresentationPlan(parentDeck, productionPlan)
assert.equal(revisedDeck.slides.length, 1)
assert.equal(revisedDeck.slides[0].id, 'p04')
assertDeckMatchesPresentationPlan(
  {
    ...revisedDeck,
    slides: parentDeck.slides.map((slide, index) => (index === 3 ? revisedDeck.slides[0] : slide)),
  },
  productionPlan,
)
const sourcePages = await pdfToPages(await readFile(join(root, 'deardorff-2020-article.pdf')))
for (const source of productionPlan.sources) {
  const isRights = source.uri === 'local:asset-rights.json'
  const bytes = await readFile(
    join(root, isRights ? 'asset-rights.json' : 'deardorff-2020-article.pdf'),
  )
  assert.equal(source.snapshotAttachmentId, digest(bytes))
  if (isRights) assert.ok(bytes.toString('utf8').includes(source.excerpt))
  else {
    const number = Number(/^PDF (\d+)/.exec(source.locator)?.[1])
    assert.ok(number && sourcePages[number - 1]?.includes(source.excerpt))
  }
}
for (const index of [4, 5]) {
  const page = parentDeck.slides[index]
  const chart = page.elements.find((element) => element.kind === 'chart')
  assert.ok(chart)
  assert.ok(
    checkPresentationChartData(productionPlan, page.id, [chart]).charts.every(
      (item) => item.findings.length === 0,
    ),
  )
}
const parentZip = await JSZip.loadAsync(await readFile(join(root, 'wiswork-parent-candidate.pptx')))
assert.equal(
  Object.keys(parentZip.files).filter((file) => /^ppt\/slides\/slide\d+\.xml$/.test(file)).length,
  8,
)
const parentPage4 = await parentZip.file('ppt/slides/slide4.xml')?.async('string')
assert.equal((parentPage4.match(/<p:pic>/g) ?? []).length, 2)
assert.ok(!parentPage4.includes('step-1-label'))
const revisedBytes = await readFile(join(root, 'revised-page-candidate.pptx'))
const revisedZip = await JSZip.loadAsync(revisedBytes)
assert.equal(
  Object.keys(revisedZip.files).filter((file) => /^ppt\/slides\/slide\d+\.xml$/.test(file)).length,
  1,
)
const revisedXml = await revisedZip.file('ppt/slides/slide1.xml')?.async('string')
for (const label of ['招募与前测', '编程工作坊', '三个月后访谈'])
  assert.ok(revisedXml.includes(label))
assert.equal((revisedXml.match(/prst="roundRect"/g) ?? []).length, 3)
assert.equal((revisedXml.match(/<p:pic>/g) ?? []).length, 2)
assert.deepEqual(manifest.preservedPageIds, plan.preserveUnchangedPageIds)
assert.equal(manifest.revisedPageId, 'p04')
assert.equal(manifest.parentPageDigests.length, 8)
assert.equal(manifest.revisedPageDigests.length, 8)
for (let index = 0; index < 8; index++) {
  const pageBytes = await readFile(
    join(root, `parent-pages/p${String(index + 1).padStart(2, '0')}.pptx`),
  )
  assert.equal(digest(pageBytes), manifest.parentPageDigests[index])
  assert.equal(
    manifest.revisedPageDigests[index],
    index === 3 ? digest(revisedBytes) : manifest.parentPageDigests[index],
  )
}
assert.notEqual(manifest.parentPageDigests[3], manifest.revisedPageDigests[3])
process.stdout.write(
  'PPT-P0-18 candidate verified: source-backed parent, exact seven retained page packages, one revised native page.\n',
)
