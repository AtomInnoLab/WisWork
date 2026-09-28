import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { PNG } from 'pngjs'

const root = dirname(fileURLToPath(import.meta.url))
const lines = (await readFile(join(root, 'SHA256SUMS'), 'utf8')).trim().split('\n')
assert.equal(lines.length, 17)
const hashes = new Map()
for (const line of lines) {
  const match = /^([a-f0-9]{64})  ([A-Za-z0-9./-]+)$/.exec(line)
  assert.ok(match && !match[2].includes('..') && !hashes.has(match[2]))
  const actual = createHash('sha256')
    .update(await readFile(join(root, match[2])))
    .digest('hex')
  assert.equal(actual, match[1], `changed ${match[2]}`)
  hashes.set(match[2], actual)
}
const scenario = JSON.parse(await readFile(join(root, 'scenario.json'), 'utf8'))
assert.equal(scenario.caseId, 'PPT-P0-15')
assert.equal(scenario.pageCount, 8)
assert.equal(scenario.targetPageNumber, 4)
assert.deepEqual(scenario.forbiddenPages, [1, 2, 3, 5, 6, 7, 8])
assert.equal(scenario.existingDeckEdits.length, 2)
assert.equal(scenario.existingDeckEdits[0].before, '参与者与证据路径')
assert.equal(scenario.existingDeckEdits[0].after, '参与者与研究证据')
assert.equal(scenario.existingDeckEdits[1].deltaTopInches, 0.08)
assert.equal(scenario.boundGeneration.expectedTitles.length, 8)
assert.equal(scenario.boundGeneration.mustGenerateAndImportWithWisWork, true)
assert.equal(scenario.boundGeneration.candidateFile, 'wiswork-generated-candidate.pptx')
assert.equal(scenario.boundGeneration.requireParentArtifactAndHostImportReceipt, true)
assert.ok(scenario.staleConflict && scenario.requiredEvidence.length >= 6)
const expectedFiles = [
  scenario.source,
  scenario.existingDeck,
  scenario.boundGeneration.candidateFile,
  'scenario.json',
  'asset-rights.json',
]
const rights = JSON.parse(await readFile(join(root, 'asset-rights.json'), 'utf8'))
assert.equal(rights.assets.length, 12)
for (const [index, asset] of rights.assets.entries()) {
  const name = `images/schematic-${String(index + 1).padStart(2, '0')}.png`
  assert.equal(asset.file, name)
  assert.equal(asset.sha256, hashes.get(name))
  assert.ok(asset.usePermission && asset.attribution)
  const image = PNG.sync.read(await readFile(join(root, name)))
  assert.deepEqual([image.width, image.height], [960, 540])
  expectedFiles.push(name)
}
assert.equal([...hashes.keys()].sort().join(','), expectedFiles.sort().join(','))
const source = await readFile(join(root, scenario.source))
assert.ok(source.equals(await readFile(join(root, '../PPT-P0-01/deardorff-2020-article.pdf'))))
const zip = await JSZip.loadAsync(await readFile(join(root, scenario.existingDeck)))
const slideFiles = Object.keys(zip.files).filter((name) =>
  /^ppt\/slides\/slide\d+\.xml$/.test(name),
)
const charts = Object.keys(zip.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))
assert.equal(slideFiles.length, 8)
assert.equal(charts.length, 2)
const page = await zip.file('ppt/slides/slide4.xml')?.async('string')
assert.ok(page?.includes(`<a:t>${scenario.existingDeckEdits[0].before}</a:t>`))
const caption = [...page.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)].find(([shape]) =>
  shape.includes('自制示意图 07'),
)?.[0]
assert.ok(
  caption && /<a:xfrm><a:off x="\d+" y="\d+"\/><a:ext cx="\d+" cy="\d+"\/><\/a:xfrm>/.test(caption),
)
const chartPage = await zip.file('ppt/slides/slide5.xml')?.async('string')
assert.ok(chartPage?.includes('<p:graphicFrame>'))
const generated = await JSZip.loadAsync(
  await readFile(join(root, scenario.boundGeneration.candidateFile)),
)
assert.equal(
  Object.keys(generated.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length,
  8,
)
assert.equal(
  Object.keys(generated.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length,
  2,
)
for (let i = 0; i < 8; i++) {
  const xml = await generated.file(`ppt/slides/slide${i + 1}.xml`)?.async('string')
  assert.ok(xml?.includes(`<a:t>${scenario.boundGeneration.expectedTitles[i]}</a:t>`))
}
assert.ok(
  (await generated.file('ppt/slides/slide4.xml')?.async('string'))?.includes(
    '招募、工作坊、三个月后访谈',
  ),
)
console.log(
  'PPT-P0-15: seventeen inputs, two eight-page native decks, four charts and exact edit targets verified',
)
