import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { checkPresentationChartData } from '@wiswork/pptx-engine/presentation-chart-data'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'
import { pdfToPages } from '../../../../packages/file-parse/src/pdf.ts'

const root = dirname(fileURLToPath(import.meta.url))
const lines = (await readFile(join(root, 'SHA256SUMS'), 'utf8')).trim().split('\n')
assert.equal(lines.length, 19)
const hashes = new Map()
for (const line of lines) {
  const match = /^([a-f0-9]{64}) {2}([A-Za-z0-9./-]+)$/.exec(line)
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
  'reference-plan.json',
  'reference-deck.json',
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
const plan = parsePresentationPlan(
  JSON.parse(await readFile(join(root, 'reference-plan.json'), 'utf8')),
)
const deck = JSON.parse(await readFile(join(root, 'reference-deck.json'), 'utf8'))
assertDeckMatchesPresentationPlan(deck, plan)
assert.equal(plan.domain, 'science')
assert.equal(plan.slides.length, 8)
assert.equal(plan.sources.length, 6)
assert.equal(plan.claims.length, 5)
assert.ok(
  plan.claims.every(
    (claim) =>
      claim.reviewStatus === 'needs_review' && claim.professionalContext?.domain === 'science',
  ),
)
assert.deepEqual(plan.slides[3].requiredAssets, ['participant-illustration'])
const paperPages = await pdfToPages(source)
assert.ok(paperPages[0].includes('Published: July 8, 2020'))
const rightsBytes = await readFile(join(root, 'asset-rights.json'))
for (const item of plan.sources) {
  const bytes = item.uri.startsWith('local:') ? rightsBytes : source
  assert.equal(item.snapshotAttachmentId, createHash('sha256').update(bytes).digest('hex'))
  if (item.uri.startsWith('local:')) assert.ok(bytes.toString('utf8').includes(item.excerpt))
  else {
    assert.equal(item.uri, 'https://doi.org/10.1371/journal.pone.0230697')
    assert.equal(item.asOf, '2020-07-08')
    const page = Number(/^PDF (\d+)/.exec(item.locator)?.[1])
    assert.ok([1, 5, 9].includes(page))
    assert.ok(paperPages[page - 1].includes(item.excerpt))
  }
}
const imageAsset = deck.assets.find((asset) => asset.id === 'participant-illustration')
assert.ok(imageAsset)
assert.equal(
  createHash('sha256').update(Buffer.from(imageAsset.base64, 'base64')).digest('hex'),
  rights.assets[6].sha256,
)
assert.equal(imageAsset.license, 'owned')
assert.match(
  deck.slides[3].elements.find((element) => element.id === 'participant-image')?.altText,
  /非研究测量数据/,
)
for (const index of [4, 5]) {
  const slide = deck.slides[index]
  const chart = slide.elements.find((element) => element.kind === 'chart')
  assert.ok(chart)
  const checked = checkPresentationChartData(plan, slide.id, [chart])
  assert.ok(checked.charts.every((entry) => entry.findings.length === 0))
}
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
const generatedPage = await generated.file('ppt/slides/slide4.xml')?.async('string')
assert.ok(generatedPage?.includes('招募、工作坊、三个月后访谈'))
assert.ok(generatedPage?.includes('<p:pic>'))
assert.ok(generatedPage?.includes('participant-image'))
assert.ok(
  generatedPage?.includes('descr="自制示意图 07：招募、工作坊与三个月后访谈；非研究测量数据"'),
)
const generatedImage = await generated.file('ppt/media/image-4-1.png')?.async('nodebuffer')
assert.ok(generatedImage)
assert.equal(createHash('sha256').update(generatedImage).digest('hex'), rights.assets[6].sha256)
for (const [chartIndex, expected] of [
  [1, ['六项清单均值', '前测', '三个月后', '1.6', '2.2']],
  [2, ['开源软件使用人数', '前测 n=14', '后测 n=12', '7', '10']],
]) {
  const chart = await generated.file(`ppt/charts/chart${chartIndex}.xml`)?.async('string')
  assert.ok(chart)
  assert.deepEqual(
    [...chart.matchAll(/<c:v>([^<]+)<\/c:v>/g)].map((match) => match[1]),
    expected,
  )
  const workbook = await generated
    .file(`ppt/embeddings/Microsoft_Excel_Worksheet${chartIndex}.xlsx`)
    ?.async('nodebuffer')
  assert.ok(workbook)
  const sheet = await (
    await JSZip.loadAsync(workbook)
  )
    .file('xl/worksheets/sheet1.xml')
    ?.async('string')
  assert.ok(sheet)
  for (const [row, value] of expected.slice(-2).entries())
    assert.ok(sheet.includes(`<c r="B${row + 2}"><v>${value}</v></c>`))
}
console.log(
  'PPT-P0-15: nineteen frozen files, source plan, two eight-page native decks, chart workbooks and exact edit targets verified',
)
